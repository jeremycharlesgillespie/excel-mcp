import { PythonShell, Options } from 'python-shell';
import path from 'path';
import { promises as fs } from 'fs';
import { ToolResult, PythonBridgeParams } from '../types/index.js';

export class PythonBridge {
  private pythonPath: string;
  private scriptsPath: string;

  constructor() {
    this.pythonPath = 'python'; // Assumes python is in PATH
    this.scriptsPath = path.join(__dirname, '..', 'python');
  }

  async callPythonFunction(params: PythonBridgeParams): Promise<ToolResult> {
    try {
      const { module, function: functionName, args = [], kwargs = {} } = params;
      
      // Create a temporary script to call the function
      const scriptContent = this.generateCallScript(module, functionName, args, kwargs);
      const tempScriptPath = path.join(this.scriptsPath, 'temp_call.py');
      
      await fs.writeFile(tempScriptPath, scriptContent);
      
      const options: Options = {
        mode: 'text',
        pythonPath: this.pythonPath,
        scriptPath: this.scriptsPath,
        args: []
      };
      
      const results = await PythonShell.run('temp_call.py', options);
      
      // Clean up temp file
      await fs.unlink(tempScriptPath).catch(() => {}); // Ignore errors
      
      // Parse JSON result
      const result = JSON.parse(results.join(''));
      
      return {
        success: true,
        data: result
      };
      
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : String(error)
      };
    }
  }

  private generateCallScript(moduleName: string, functionName: string, args: any[], kwargs: Record<string, any>): string {
    const argsJson = JSON.stringify(args);
    const kwargsJson = JSON.stringify(kwargs);
    
    return `
import json
import sys
import os
from datetime import datetime, date
import pandas as pd
import numpy as np

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    # Import the module
    module = __import__('${moduleName}')
    
    # Parse arguments
    args = json.loads('${argsJson}')
    kwargs = json.loads('${kwargsJson}')
    
    # Convert date strings to date objects if needed
    def convert_dates(obj):
        if isinstance(obj, dict):
            for key, value in obj.items():
                if isinstance(value, str) and len(value) == 10 and value.count('-') == 2:
                    try:
                        obj[key] = datetime.strptime(value, '%Y-%m-%d').date()
                    except ValueError:
                        pass
                elif isinstance(value, (dict, list)):
                    obj[key] = convert_dates(value)
        elif isinstance(obj, list):
            return [convert_dates(item) for item in obj]
        return obj
    
    args = convert_dates(args)
    kwargs = convert_dates(kwargs)
    
    # Get the function
    if '.' in '${functionName}':
        # Class method call
        class_name, method_name = '${functionName}'.split('.', 1)
        cls = getattr(module, class_name)
        if method_name == '__init__':
            result = cls(*args, **kwargs)
        else:
            func = getattr(cls, method_name)
            if hasattr(func, '__self__'):  # Instance method
                instance = cls()
                result = func(*args, **kwargs)
            else:  # Static method
                result = func(*args, **kwargs)
    else:
        # Direct function call
        func = getattr(module, '${functionName}')
        result = func(*args, **kwargs)
    
    # Convert result to JSON-serializable format
    def make_serializable(obj):
        if isinstance(obj, pd.DataFrame):
            return obj.to_dict('records')
        elif isinstance(obj, pd.Series):
            return obj.to_dict()
        elif isinstance(obj, (date, datetime)):
            return obj.isoformat()
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, dict):
            return {k: make_serializable(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [make_serializable(item) for item in obj]
        elif hasattr(obj, '__dict__'):
            return make_serializable(obj.__dict__)
        return obj
    
    output = make_serializable(result)
    print(json.dumps(output, default=str))
    
except Exception as e:
    error_result = {
        'error': str(e),
        'type': type(e).__name__
    }
    print(json.dumps(error_result))
`;
  }

  async testConnection(): Promise<ToolResult> {
    try {
      const testScript = `
import json
import sys
print(json.dumps({
    "python_version": sys.version,
    "modules_available": {
        "pandas": True,
        "numpy": True,
        "openpyxl": True
    }
}))
`;
      
      const tempPath = path.join(this.scriptsPath, 'test_connection.py');
      await fs.writeFile(tempPath, testScript);
      
      const options: Options = {
        mode: 'text',
        pythonPath: this.pythonPath,
        scriptPath: this.scriptsPath
      };
      
      const results = await PythonShell.run('test_connection.py', options);
      await fs.unlink(tempPath).catch(() => {});
      
      return {
        success: true,
        data: JSON.parse(results.join(''))
      };
      
    } catch (error) {
      return {
        success: false,
        error: `Python bridge connection failed: ${error instanceof Error ? error.message : String(error)}`
      };
    }
  }

  async installRequirements(): Promise<ToolResult> {
    try {
      const requirementsPath = path.join(__dirname, '..', '..', 'requirements.txt');
      
      const options: Options = {
        mode: 'text',
        pythonPath: this.pythonPath,
        args: ['-m', 'pip', 'install', '-r', requirementsPath]
      };
      
      const results = await PythonShell.run('', options);
      
      return {
        success: true,
        data: results.join('\n')
      };
      
    } catch (error) {
      return {
        success: false,
        error: `Failed to install requirements: ${error instanceof Error ? error.message : String(error)}`
      };
    }
  }
}