"""
Complete Utilities and Integration Helpers for MOPS Sheets Uploader v1.1.1
Create this new file: utils.py
"""

import os
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional
import logging

logger = logging.getLogger(__name__)

class ComponentLoader:
    """Safe component loader with fallbacks for v1.1.1"""
    
    @staticmethod
    def safe_import_component(module_name: str, class_name: str, config=None):
        """Safely import components with fallback handling"""
        try:
            module = __import__(f".{module_name}", fromlist=[class_name], package="mops_sheets_uploader")
            component_class = getattr(module, class_name)
            return component_class(config) if config else component_class
        except ImportError as e:
            logger.warning(f"Could not import {class_name} from {module_name}: {e}")
            return None
        except Exception as e:
            logger.error(f"Error initializing {class_name}: {e}")
            return None

def init_components_safely(config) -> Dict[str, Any]:
    """Initialize all components safely with enhanced error handling"""
    loader = ComponentLoader()
    
    components = {
        'pdf_scanner': loader.safe_import_component('pdf_scanner', 'PDFScanner', config),
        'stock_loader': loader.safe_import_component('stock_data_loader', 'StockDataLoader', config),
        'matrix_builder': loader.safe_import_component('matrix_builder', 'MatrixBuilder', config),
        'upload_manager': loader.safe_import_component('sheets_connector', 'SheetsUploadManager', config),
        'analyzer': loader.safe_import_component('report_analyzer', 'ReportAnalyzer', config)
    }
    
    # Check which components loaded successfully
    loaded_components = [name for name, component in components.items() if component is not None]
    failed_components = [name for name, component in components.items() if component is None]
    
    if failed_components:
        logger.warning(f"⚠️ Some components failed to load: {failed_components}")
        logger.info(f"✅ Successfully loaded: {loaded_components}")
    else:
        logger.info(f"✅ All components loaded successfully: {loaded_components}")
    
    return components

def format_multiple_types_display(report_types: set, 
                                 separator: str = "/",
                                 max_types: int = 5,
                                 use_priority_sort: bool = True) -> str:
    """
    Format multiple report types for display (v1.1.1)
    
    Args:
        report_types: Set of report type strings
        separator: Separator between types
        max_types: Maximum types to show before truncation
        use_priority_sort: Whether to sort by priority
        
    Returns:
        Formatted string for display
    """
    if not report_types:
        return '-'
    
    # Convert to list and sort
    type_list = list(report_types)
    
    if use_priority_sort:
        try:
            from .models import REPORT_TYPE_PRIORITY
            type_list.sort(key=lambda x: (REPORT_TYPE_PRIORITY.get(x, 9), x))
        except ImportError:
            type_list.sort()
    else:
        type_list.sort()
    
    # Apply truncation if needed
    if len(type_list) <= max_types:
        return separator.join(type_list)
    else:
        visible_types = type_list[:max_types-1]
        remaining_count = len(type_list) - len(visible_types)
        return separator.join(visible_types) + f"+{remaining_count}"

def create_categorized_display(report_types: set,
                              type_separator: str = "/",
                              category_separator: str = " → ") -> str:
    """
    Create categorized display for multiple report types (v1.1.1)
    """
    try:
        from .models import REPORT_TYPE_CATEGORIES
    except ImportError:
        # Fallback categories
        REPORT_TYPE_CATEGORIES = {
            'individual': ['A12', 'A13'],
            'consolidated': ['AI1', 'A1L'],
            'generic': ['A10', 'A11'],
            'english': ['AIA', 'AE2'],
            'other': []
        }
    
    from collections import defaultdict
    
    # Group by category
    category_groups = defaultdict(list)
    
    for report_type in report_types:
        category_found = False
        for category, types in REPORT_TYPE_CATEGORIES.items():
            if report_type in types:
                category_groups[category].append(report_type)
                category_found = True
                break
        if not category_found:
            category_groups['other'].append(report_type)
    
    # Build display string
    category_order = ['individual', 'consolidated', 'generic', 'other']
    display_parts = []
    
    for category in category_order:
        if category in category_groups:
            sorted_types = sorted(category_groups[category])
            display_parts.append(type_separator.join(sorted_types))
    
    return category_separator.join(display_parts)

def validate_and_fix_config(config) -> Any:
    """Validate configuration and auto-fix common issues"""
    fixes_applied = []
    
    # Fix font sizes
    if hasattr(config, 'font_size') and (config.font_size < 8 or config.font_size > 72):
        config.font_size = 14
        fixes_applied.append("font_size reset to 14")
    
    if hasattr(config, 'header_font_size') and (config.header_font_size < 8 or config.header_font_size > 72):
        config.header_font_size = 14
        fixes_applied.append("header_font_size reset to 14")
    
    # Fix multiple type settings
    if hasattr(config, 'max_display_types') and (config.max_display_types < 1 or config.max_display_types > 20):
        config.max_display_types = 5
        fixes_applied.append("max_display_types reset to 5")
    
    if hasattr(config, 'report_type_separator') and len(config.report_type_separator) > 5:
        config.report_type_separator = "/"
        fixes_applied.append("report_type_separator reset to '/'")
    
    # Ensure required directories exist
    os.makedirs(config.downloads_dir, exist_ok=True)
    os.makedirs(config.output_dir, exist_ok=True)
    os.makedirs("logs", exist_ok=True)
    
    if fixes_applied:
        logger.info(f"⚙️ Configuration auto-fixes applied: {', '.join(fixes_applied)}")
    
    return config

def setup_v1_1_1_environment() -> Dict[str, Any]:
    """Setup environment for v1.1.1 features"""
    import os
    from pathlib import Path
    
    setup_result = {
        'directories_created': [],
        'environment_variables_set': [],
        'status': 'success',
        'errors': []
    }
    
    try:
        # Ensure required directories exist
        directories = ['downloads', 'logs', 'data/reports', 'data/exports']
        for directory in directories:
            path = Path(directory)
            if not path.exists():
                path.mkdir(parents=True, exist_ok=True)
                setup_result['directories_created'].append(directory)
        
        # Set v1.1.1 environment defaults if not set
        env_defaults = {
            'MOPS_SHOW_ALL_TYPES': 'true',
            'MOPS_TYPE_SEPARATOR': '/',
            'MOPS_FONT_SIZE': '14',
            'MOPS_HEADER_FONT_SIZE': '14',
            'MOPS_BOLD_HEADERS': 'true',
            'MOPS_BOLD_COMPANY_INFO': 'true',
            'MOPS_HIGHLIGHT_MULTIPLE': 'true',
            'MOPS_WORKSHEET_NAME': 'MOPS下載狀態'
        }
        
        for key, value in env_defaults.items():
            if key not in os.environ:
                os.environ[key] = value
                setup_result['environment_variables_set'].append(key)
        
        logger.info(f"📝 Environment setup completed successfully")
        if setup_result['directories_created']:
            logger.info(f"   Created directories: {', '.join(setup_result['directories_created'])}")
        if setup_result['environment_variables_set']:
            logger.info(f"   Set environment variables: {', '.join(setup_result['environment_variables_set'])}")
        
    except Exception as e:
        setup_result['status'] = 'error'
        setup_result['errors'].append(str(e))
        logger.error(f"Environment setup failed: {e}")
    
    return setup_result

def test_v1_1_1_features() -> Dict[str, Any]:
    """Test v1.1.1 specific features"""
    test_results = {
        'font_presets': 'not_tested',
        'multiple_types': 'not_tested', 
        'config_loading': 'not_tested',
        'models_import': 'not_tested',
        'overall_status': 'unknown'
    }
    
    # Test font configuration
    try:
        from .models import create_font_config_preset
        preset = create_font_config_preset('large')
        test_results['font_presets'] = preset is not None and preset.get('font_size') == 14
    except Exception as e:
        test_results['font_presets'] = f"Failed: {e}"
    
    # Test multiple types display
    try:
        from .models import MatrixCell, PDFFile
        from datetime import datetime
        
        cell = MatrixCell()
        pdf1 = PDFFile("2330", 2024, 1, "A12", "test1.pdf", "/path1", 1000, datetime.now())
        pdf2 = PDFFile("2330", 2024, 1, "A13", "test2.pdf", "/path2", 1000, datetime.now())
        
        cell.add_pdf(pdf1)
        cell.add_pdf(pdf2)
        
        display_value = cell.get_display_value(show_all_types=True)
        test_results['multiple_types'] = "A12/A13" in display_value or "A13/A12" in display_value
    except Exception as e:
        test_results['multiple_types'] = f"Failed: {e}"
    
    # Test configuration loading
    try:
        from .config import load_config
        config = load_config()
        test_results['config_loading'] = hasattr(config, 'font_size')
    except Exception as e:
        test_results['config_loading'] = f"Failed: {e}"
    
    # Test models import
    try:
        from .models import REPORT_TYPE_CATEGORIES, STATUS_MAPPING
        test_results['models_import'] = bool(REPORT_TYPE_CATEGORIES and STATUS_MAPPING)
    except Exception as e:
        test_results['models_import'] = f"Failed: {e}"
    
    # Determine overall status
    passed_tests = sum(1 for result in test_results.values() 
                      if result is True)
    total_tests = len([k for k in test_results.keys() if k != 'overall_status'])
    
    if passed_tests == total_tests:
        test_results['overall_status'] = 'all_passed'
    elif passed_tests > 0:
        test_results['overall_status'] = 'partial_pass'
    else:
        test_results['overall_status'] = 'all_failed'
    
    # Print test results
    print("🧪 v1.1.1 Feature Tests:")
    for feature, result in test_results.items():
        if feature == 'overall_status':
            continue
        
        if result is True:
            status = "✅"
        elif result is False:
            status = "❌"
        else:
            status = "⚠️"
        
        print(f"   {status} {feature}: {result}")
    
    print(f"📊 Overall Status: {test_results['overall_status']} ({passed_tests}/{total_tests} passed)")
    
    return test_results

def migrate_to_v1_1_1(old_config_path: str = None) -> bool:
    """Help migrate from older versions to v1.1.1"""
    print("🔄 Migrating to MOPS Sheets Uploader v1.1.1...")
    
    migration_success = True
    
    try:
        # Step 1: Setup environment
        print("📝 Setting up v1.1.1 environment...")
        setup_result = setup_v1_1_1_environment()
        
        if setup_result['status'] != 'success':
            print(f"⚠️ Environment setup had issues: {setup_result['errors']}")
            migration_success = False
        
        # Step 2: Test features
        print("🧪 Testing v1.1.1 features...")
        test_results = test_v1_1_1_features()
        
        if test_results['overall_status'] == 'all_failed':
            print("❌ Critical: All feature tests failed")
            migration_success = False
        elif test_results['overall_status'] == 'partial_pass':
            print("⚠️ Warning: Some feature tests failed")
        
        # Step 3: Migration summary
        print(f"📊 Migration Summary:")
        print(f"   • Environment Setup: {'✅' if setup_result['status'] == 'success' else '❌'}")
        print(f"   • Feature Tests: {'✅' if test_results['overall_status'] == 'all_passed' else '⚠️' if test_results['overall_status'] == 'partial_pass' else '❌'}")
        print(f"   • Font Support: {'✅' if test_results['font_presets'] is True else '❌'}")
        print(f"   • Multiple Types: {'✅' if test_results['multiple_types'] is True else '❌'}")
        print(f"   • Configuration: {'✅' if test_results['config_loading'] is True else '❌'}")
        
        if migration_success and test_results['overall_status'] == 'all_passed':
            print("🎉 Migration to v1.1.1 completed successfully!")
            print("💡 You can now use enhanced features:")
            print("   • Font presets: font_preset='large'")
            print("   • Multiple types: show_all_report_types=True")
            print("   • 14pt professional styling by default")
            return True
        else:
            print("⚠️ Migration completed with warnings. Check error messages above.")
            return False
            
    except Exception as e:
        print(f"❌ Migration failed: {e}")
        return False

def get_version_info() -> Dict[str, Any]:
    """Get comprehensive version information"""
    try:
        from . import __version__, __author__, __description__
    except ImportError:
        __version__ = "1.1.1"
        __author__ = "MOPS Analysis Team"
        __description__ = "MOPS Sheets Uploader"
    
    version_info = {
        'version': __version__,
        'author': __author__,
        'description': __description__,
        'python_version': f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        'python_compatible': sys.version_info >= (3, 9),
        'features': {
            'multiple_report_types': True,
            'font_configuration': True,
            'enhanced_analytics': True,
            'professional_styling': True
        }
    }
    
    # Test feature availability
    try:
        from .models import create_font_config_preset
        version_info['features']['font_presets_available'] = True
    except ImportError:
        version_info['features']['font_presets_available'] = False
    
    try:
        from .models import MatrixCell
        test_cell = MatrixCell()
        version_info['features']['multiple_types_working'] = hasattr(test_cell, 'get_display_value')
    except ImportError:
        version_info['features']['multiple_types_working'] = False
    
    return version_info

def print_system_info():
    """Print comprehensive system information"""
    info = get_version_info()
    
    print("🔍 System Information:")
    print("=" * 40)
    print(f"MOPS Sheets Uploader: v{info['version']}")
    print(f"Python Version: {info['python_version']}")
    print(f"Python Compatible: {'✅' if info['python_compatible'] else '❌'}")
    print()
    print("🎯 v1.1.1 Features:")
    for feature, available in info['features'].items():
        status = "✅" if available else "❌"
        feature_name = feature.replace('_', ' ').title()
        print(f"   {status} {feature_name}")
    
    print()
    print("📁 Directory Status:")
    required_dirs = ['downloads', 'logs', 'data/reports']
    for directory in required_dirs:
        exists = Path(directory).exists()
        status = "✅" if exists else "❌"
        print(f"   {status} {directory}")
    
    print()
    print("🔧 Environment Variables:")
    required_env = ['GOOGLE_SHEETS_CREDENTIALS', 'GOOGLE_SHEET_ID']
    optional_env = ['MOPS_FONT_SIZE', 'MOPS_SHOW_ALL_TYPES', 'MOPS_WORKSHEET_NAME']
    
    for env_var in required_env:
        exists = bool(os.getenv(env_var))
        status = "✅" if exists else "❌"
        print(f"   {status} {env_var} (required)")
    
    for env_var in optional_env:
        exists = bool(os.getenv(env_var))
        value = os.getenv(env_var, 'not set')
        status = "✅" if exists else "○"
        print(f"   {status} {env_var} = {value}")

def ensure_v1_1_1_compatibility():
    """Ensure backward compatibility while enabling v1.1.1 features"""
    import os
    
    # Set environment defaults for v1.1.1 if not present
    env_defaults = {
        'MOPS_SHOW_ALL_TYPES': 'true',
        'MOPS_FONT_SIZE': '14',
        'MOPS_HEADER_FONT_SIZE': '14',
        'MOPS_BOLD_HEADERS': 'true',
        'MOPS_BOLD_COMPANY_INFO': 'true'
    }
    
    set_vars = []
    for key, default_value in env_defaults.items():
        if key not in os.environ:
            os.environ[key] = default_value
            set_vars.append(key)
    
    if set_vars:
        logger.info(f"📝 Set v1.1.1 compatibility defaults: {', '.join(set_vars)}")
    
    return set_vars

class V1_1_1_Validator:
    """Validator for v1.1.1 specific features and requirements"""
    
    @staticmethod
    def validate_font_config(font_config: Dict[str, Any]) -> List[str]:
        """Validate font configuration"""
        errors = []
        
        required_fields = ['font_size', 'header_font_size', 'bold_headers', 'bold_company_info']
        for field in required_fields:
            if field not in font_config:
                errors.append(f"Missing font config field: {field}")
        
        if 'font_size' in font_config:
            size = font_config['font_size']
            if not isinstance(size, int) or size < 8 or size > 72:
                errors.append("font_size must be integer between 8-72")
        
        if 'header_font_size' in font_config:
            size = font_config['header_font_size']
            if not isinstance(size, int) or size < 8 or size > 72:
                errors.append("header_font_size must be integer between 8-72")
        
        return errors
    
    @staticmethod
    def validate_multiple_type_config(config) -> List[str]:
        """Validate multiple report type configuration"""
        errors = []
        
        if hasattr(config, 'max_display_types'):
            if config.max_display_types < 1 or config.max_display_types > 20:
                errors.append("max_display_types must be between 1-20")
        
        if hasattr(config, 'report_type_separator'):
            if len(config.report_type_separator) > 5:
                errors.append("report_type_separator too long (max 5 chars)")
        
        return errors
    
    @staticmethod
    def validate_full_v1_1_1_setup() -> Dict[str, Any]:
        """Comprehensive v1.1.1 setup validation"""
        validation = {
            'directories': {},
            'environment': {},
            'features': {},
            'overall_status': True
        }
        
        # Directory validation
        required_dirs = ['downloads', 'logs', 'data/reports']
        for directory in required_dirs:
            exists = Path(directory).exists()
            validation['directories'][directory] = exists
            if not exists:
                validation['overall_status'] = False
        
        # Environment validation
        required_env = ['GOOGLE_SHEETS_CREDENTIALS', 'GOOGLE_SHEET_ID']
        for env_var in required_env:
            exists = bool(os.getenv(env_var))
            validation['environment'][env_var] = exists
            if not exists:
                validation['overall_status'] = False
        
        # Feature validation
        feature_tests = test_v1_1_1_features()
        validation['features'] = feature_tests
        if feature_tests['overall_status'] == 'all_failed':
            validation['overall_status'] = False
        
        return validation