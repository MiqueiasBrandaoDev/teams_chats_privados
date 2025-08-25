#!/usr/bin/env python3

import sys
from device_auth import DeviceCodeAuthenticator

def test_device_authentication():
    print("🔐 Testando autenticação via Device Code...")
    print("💡 Método mais confiável para WSL/Linux")
    
    try:
        authenticator = DeviceCodeAuthenticator()
        success, message = authenticator.test_connection()
        
        if success:
            print(f"✅ {message}")
            print("🚀 Pronto para exportar chats privados!")
            return True
        else:
            print(f"❌ {message}")
            return False
            
    except Exception as e:
        print(f"❌ Erro na autenticação: {str(e)}")
        return False

if __name__ == "__main__":
    success = test_device_authentication()
    sys.exit(0 if success else 1)