"""
Test script for Office Pro Simplified API
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def test_import():
    """Test module import."""
    print("=" * 60)
    print("Test 1: Module Import")
    print("=" * 60)
    
    try:
        from office_pro import (
            QuickGenerator,
            generate_contract,
            generate_report,
            generate_excel,
            list_templates,
            num_to_chinese,
            check_dependencies,
            get_version_info,
            ExcelChartBuilder,
            WordStyleBuilder,
            create_chart,
            create_styled_document,
        )
        print("✓ All imports successful")
        return True
    except ImportError as e:
        print(f"✗ Import failed: {e}")
        return False


def test_dependencies():
    """Test dependency checking."""
    print("\n" + "=" * 60)
    print("Test 2: Dependency Check")
    print("=" * 60)
    
    try:
        import office_pro
        deps = office_pro.check_dependencies()
        print("Dependencies status:")
        for dep, status in deps.items():
            symbol = "✓" if status else "✗"
            print(f"  {symbol} {dep}: {'Available' if status else 'Not installed'}")
        
        info = office_pro.get_version_info()
        print(f"\nVersion: {info['version']}")
        return True
    except Exception as e:
        print(f"✗ Dependency check failed: {e}")
        return False


def test_num_to_chinese():
    """Test number to Chinese conversion."""
    print("\n" + "=" * 60)
    print("Test 3: Number to Chinese")
    print("=" * 60)
    
    try:
        from office_pro import num_to_chinese
        
        test_cases = [
            (500, "伍佰"),
            (1000, "壹仟"),
            (15000, "壹万伍仟"),
            (3500, "叁仟伍佰"),
        ]
        
        all_passed = True
        for num, expected in test_cases:
            result = num_to_chinese(num)
            status = "✓" if result == expected else "✗"
            print(f"  {status} {num} -> {result}")
            if result != expected:
                all_passed = False
        
        return all_passed
    except Exception as e:
        print(f"✗ Test failed: {e}")
        return False


def test_list_templates():
    """Test listing templates."""
    print("\n" + "=" * 60)
    print("Test 4: List Templates")
    print("=" * 60)
    
    try:
        from office_pro import list_templates
        
        all_templates = list_templates()
        print("Available templates:")
        for category, templates in all_templates.items():
            print(f"  {category}: {len(templates)} templates")
        
        return True
    except Exception as e:
        print(f"✗ Test failed: {e}")
        return False


def test_generate_contract():
    """Test generating a contract."""
    print("\n" + "=" * 60)
    print("Test 5: Generate Contract")
    print("=" * 60)
    
    try:
        from office_pro import generate_contract
        
        try:
            import docx
        except ImportError:
            print("⚠ python-docx not installed, skipping")
            return True
        
        print("Generating parking lease contract...")
        result = generate_contract(
            'parking_lease',
            party_a='张三',
            party_b='李四',
            location='XX小区地下停车场',
            space_number='A-088',
            monthly_rent=500,
            start_date='2024-01-01',
            end_date='2024-12-31',
            output='test_contract.docx'
        )
        
        if os.path.exists(result):
            size = os.path.getsize(result)
            print(f"✓ Contract generated: {result} ({size} bytes)")
            os.remove(result)
            return True
        else:
            print(f"✗ File not found: {result}")
            return False
            
    except Exception as e:
        print(f"✗ Test failed: {e}")
        return False


def test_generate_excel():
    """Test generating an Excel file."""
    print("\n" + "=" * 60)
    print("Test 6: Generate Excel")
    print("=" * 60)
    
    try:
        from office_pro import generate_excel
        
        try:
            import openpyxl
        except ImportError:
            print("⚠ openpyxl not installed, skipping")
            return True
        
        print("Generating financial report...")
        result = generate_excel(
            'financial_report',
            company_name='测试公司',
            output='test_excel.xlsx'
        )
        
        if os.path.exists(result):
            size = os.path.getsize(result)
            print(f"✓ Excel generated: {result} ({size} bytes)")
            os.remove(result)
            return True
        else:
            print(f"✗ File not found: {result}")
            return False
            
    except Exception as e:
        print(f"✗ Test failed: {e}")
        return False


def main():
    """Run all tests."""
    print("\n" + "=" * 60)
    print("Office Pro - Test Suite")
    print("=" * 60)
    
    tests = [
        ("Import", test_import),
        ("Dependencies", test_dependencies),
        ("NumToChinese", test_num_to_chinese),
        ("ListTemplates", test_list_templates),
        ("GenerateContract", test_generate_contract),
        ("GenerateExcel", test_generate_excel),
    ]
    
    results = []
    for name, test_func in tests:
        result = test_func()
        results.append((name, result))
    
    print("\n" + "=" * 60)
    print("Summary")
    print("=" * 60)
    
    passed = sum(1 for _, r in results if r)
    total = len(results)
    
    for name, result in results:
        status = "✓" if result else "✗"
        print(f"  {status} {name}")
    
    print(f"\nTotal: {passed}/{total} passed")
    
    return 0 if passed == total else 1


if __name__ == '__main__':
    sys.exit(main())
