import os

# Version of the package
__version__ = '1.0.0'

# Conditional imports
if os.getenv('ENV') == 'production':
    from .Utils import extract_outlook_data as ProductionClass
else:
    from .Utils import extract_outlook_data as ExtractData

# Public API of the package
__all__ = ['ProductionClass', 'ExtractData']