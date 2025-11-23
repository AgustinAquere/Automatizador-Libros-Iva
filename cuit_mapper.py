"""
Maneja el mapeo entre CUIT y nombres de clientes
"""

import json
import os

CUIT_MAP_FILE = 'cuit_mapping.json'

class CUITMapper:
    """Maneja el mapeo CUIT -> Cliente"""
    
    def __init__(self):
        self.mapping = self.load_mapping()
    
    def load_mapping(self) -> dict:
        """Carga el mapeo desde el archivo JSON"""
        if os.path.exists(CUIT_MAP_FILE):
            with open(CUIT_MAP_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}
    
    def save_mapping(self):
        """Guarda el mapeo en el archivo JSON"""
        with open(CUIT_MAP_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.mapping, f, indent=2, ensure_ascii=False)
    
    def add_client(self, cuit: str, client_name: str):
        """Agrega un cliente al mapeo"""
        self.mapping[cuit] = client_name
        self.save_mapping()
    
    def get_client_by_cuit(self, cuit: str) -> str:
        """Obtiene el nombre del cliente por CUIT"""
        return self.mapping.get(cuit, None)
    
    def get_cuit_by_client(self, client_name: str) -> str:
        """Obtiene el CUIT por nombre de cliente"""
        for cuit, name in self.mapping.items():
            if name == client_name:
                return cuit
        return None
    
    def client_exists(self, cuit: str) -> bool:
        """Verifica si un CUIT ya existe"""
        return cuit in self.mapping
    
    def get_all_clients(self) -> dict:
        """Retorna todos los clientes"""
        return self.mapping.copy()
