# interfaz_base.py
from abc import ABC, abstractmethod

class InterfazUsuario(ABC):
    """Clase abstracta que define el contrato para cualquier interfaz de usuario"""
    
    @abstractmethod
    def ejecutar(self, gestor):
        """
        Método principal que inicia la interfaz
        
        Args:
            gestor: Instancia de GestorPlantillas con la lógica de la aplicación
        """
        pass
    
    @abstractmethod
    def mostrar_popup_comandos(self, filtro, x=None, y=None):
        """Muestra un popup con los comandos que coinciden con el filtro"""
        pass
    
    @abstractmethod
    def actualizar_popup_comandos(self, filtro):
        """Actualiza un popup existente con un nuevo filtro"""
        pass
    
    @abstractmethod
    def mostrar_mensaje(self, titulo, mensaje, tipo="info"):
        """Muestra un mensaje al usuario"""
        pass
    
    @abstractmethod
    def preguntar(self, titulo, mensaje):
        """Pregunta al usuario y devuelve True/False"""
        pass
    
    @abstractmethod
    def seleccionar_carpeta(self, titulo, directorio_inicial):
        """Abre un diálogo para seleccionar carpeta"""
        pass