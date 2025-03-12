import json
import os
from datetime import datetime

CONFIG_FILE = 'config.json'
HISTORY_LIMIT = 10

class Config:
    def __init__(self):
        self.config_path = os.path.join(os.path.dirname(__file__), '..', CONFIG_FILE)
        self.data = self._load_config()
    
    def _load_config(self):
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return self._get_default_config()
        return self._get_default_config()
    
    def _get_default_config(self):
        return {
            'recent_inputs': {},
            'recent_files': [],
            'common_values': {
                'seller_email': [],
                'seller_name': [],
                'store_link': []
            }
        }
    
    def save_config(self):
        os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
    
    def update_recent_inputs(self, template_name, form_data):
        if 'recent_inputs' not in self.data:
            self.data['recent_inputs'] = {}
        self.data['recent_inputs'][template_name] = form_data
        self.save_config()
    
    def get_recent_inputs(self, template_name):
        return self.data.get('recent_inputs', {}).get(template_name, {})
    
    def add_recent_file(self, file_path):
        if 'recent_files' not in self.data:
            self.data['recent_files'] = []
        
        # 确保路径是相对路径
        rel_path = os.path.relpath(file_path, os.path.dirname(self.config_path))
        
        # 移除已存在的相同文件
        self.data['recent_files'] = [f for f in self.data['recent_files'] if f != rel_path]
        
        # 添加到开头
        self.data['recent_files'].insert(0, rel_path)
        
        # 限制数量
        self.data['recent_files'] = self.data['recent_files'][:HISTORY_LIMIT]
        
        self.save_config()
    
    def get_recent_files(self):
        base_dir = os.path.dirname(self.config_path)
        return [os.path.join(base_dir, f) for f in self.data.get('recent_files', [])]
    
    def update_common_values(self, field_name, value):
        if 'common_values' not in self.data:
            self.data['common_values'] = {}
        if field_name not in self.data['common_values']:
            self.data['common_values'][field_name] = []
        
        # 对于多值字段（如store_link），分别添加
        if field_name == 'store_link':
            values = [v.strip() for v in value.split(',')]
        else:
            values = [value]
        
        for val in values:
            if val and val not in self.data['common_values'][field_name]:
                self.data['common_values'][field_name].insert(0, val)
                # 限制数量
                self.data['common_values'][field_name] = \
                    self.data['common_values'][field_name][:HISTORY_LIMIT]
        
        self.save_config()
    
    def get_common_values(self, field_name):
        return self.data.get('common_values', {}).get(field_name, [])