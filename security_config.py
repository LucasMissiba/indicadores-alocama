"""
Configurações de segurança para o Dashboard ALOCAMA
"""

# Configurações de segurança
SECURITY_CONFIG = {
    # Limites de arquivo
    'MAX_FILE_SIZE': 50 * 1024 * 1024,  # 50MB
    'MAX_ROWS': 100000,  # Máximo de linhas por DataFrame
    'MAX_COLUMNS': 100,  # Máximo de colunas por DataFrame
    'MAX_FILES': 1000,   # Máximo de arquivos para processar
    
    # Extensões permitidas
    'ALLOWED_EXTENSIONS': ['.xlsx', '.xls'],
    
    # Padrões suspeitos para detectar
    'SUSPICIOUS_PATTERNS': [
        '<script',
        'javascript:',
        'onload=',
        'onerror=',
        'eval(',
        'document.cookie',
        'window.location'
    ],
    
    # Configurações de logging
    'LOG_LEVEL': 'INFO',
    'LOG_FILE': 'dashboard.log',
    'LOG_MAX_SIZE': 10 * 1024 * 1024,  # 10MB
    'LOG_BACKUP_COUNT': 5,
    
    # Configurações de sessão
    'SESSION_TIMEOUT': 3600,  # 1 hora
    'MAX_SESSION_SIZE': 100 * 1024 * 1024,  # 100MB
}

# Configurações de performance
PERFORMANCE_CONFIG = {
    'CHUNK_SIZE': 1000,  # Processar arquivos em chunks
    'MAX_WORKERS': 4,    # Máximo de workers para processamento paralelo
    'CACHE_TTL': 300,    # 5 minutos de cache
    'MEMORY_LIMIT': 512 * 1024 * 1024,  # 512MB limite de memória
}

# Configurações de validação de dados
DATA_VALIDATION_CONFIG = {
    'REQUIRED_COLUMNS': ['item', 'quantidade', 'valor'],
    'NUMERIC_COLUMNS': ['quantidade', 'valor', 'faturamento'],
    'DATE_COLUMNS': ['data', 'mes', 'ano'],
    'MAX_STRING_LENGTH': 1000,
    'MIN_NUMERIC_VALUE': 0,
    'MAX_NUMERIC_VALUE': 999999999,
}
