"""
Test script for Office Pro Simplified API

This script tests the new simplified API for generating contracts and reports.
"""

import sys
import os

# Add the current directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_import():
    """Test that the module can be imported."""
    print("=" * 60)
    print("Test 1: Module Import")
    print("=" * 60)
    
    try:
        from office_pro import (
            QuickGenerator,
            generate_contract,
            generate_report,
            list_templates,
            num_to_chinese,
            check_dependencies,
            get_version_info,
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
        print(f"Legacy API: {'Available' if info['legacy_api_available'] else 'Not available'}")
        return True
    except Exception as e:
        print(f"✗ Dependency check failed: {e}")
        return False


def test_num_to_chinese():
    """Test number to Chinese conversion."""
    print("\n" + "=" * 60)
    print("Test 3: Number to Chinese Conversion")
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
            print(f"  {status} {num} -> {result} (expected: {expected})")
            if result != expected:
                all_passed = False
        
        return all_passed
    except Exception as e:
        print(f"✗ Test failed: {e}")
        return False


def test_list_templates():
    """Test listing available templates."""
    print("\n" + "=" * 60)
    print("Test 4: List Templates")
    print("=" * 60)
    
    try:
        from office_pro import list_templates
        
        # List all templates
        all_templates = list_templates()
        print("Available templates:")
        for category, templates in all_templates.items():
            print(f"  {category}:")
            for t in templates:
                print(f"    - {t}")
        
        # List only contracts
        contracts = list_templates('contract')
        print(f"\nContracts: {contracts.get('contracts', [])}")
        
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
        
        # Check if python-docx is available
        try:
            import docx
        except ImportError:
            print("⚠ python-docx not installed, skipping contract generation test")
            print("  Install with: pip install python-docx")
            return True
        
        # Generate a parking lease contract
        print("Generating parking lease contract...")
        result = generate_contract(
            'parking_lease',
            party_a='张三',
            party_b='李四',
            location='XX小区地下停车场B2层',
            space_number='A-088',
            monthly_rent=500,
            deposit=1000,
            start_date='2024-01-01',
            end_date='2024-12-31',
            output='test_parking_contract.docx'
        )
        
        if os.path.exists(result):
            file_size = os.path.getsize(result)
            print(f"✓ Contract generated successfully")
            print(f"  File: {result}")
            print(f"  Size: {file_size} bytes")
            
            # Clean up
            os.remove(result)
            print("  (Test file removed)")
            return True
        else:
            print(f"✗ File not found: {result}")
            return False
            
    except Exception as e:
        print(f"✗ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_generate_report():
    """Test generating a report."""
    print("\n" + "=" * 60)
    print("Test 6: Generate Report")
    print("=" * 60)
    
    try:
        from office_pro import generate_report
        
        # Check if python-docx is available
        try:
            import docx
        except ImportError:
            print("⚠ python-docx not installed, skipping report generation test")
            print("  Install with: pip install python-docx")
            return True
        
        # Generate meeting minutes
        print("Generating meeting minutes...")
        result = generate_report(
            'meeting_minutes',
            meeting_title='Q1产品规划会议',
            meeting_date='2024-03-15',
            chairperson='张三',
            secretary='李四',
            attendees='张三、李四、王五、赵六',
            agenda='1. 产品路线图回顾\n2. 新功能讨论\n3. 资源分配',
            discussion='会议讨论了Q1的产品规划，确定了核心功能...',
            decisions='1. 确定3个核心功能\n2. 分配开发资源',
            action_items='1. 张三负责功能设计 - 3月20日前\n2. 李四负责技术方案 - 3月22日前',
            output='test_meeting_minutes.docx'
        )
        
        if os.path.exists(result):
            file_size = os.path.getsize(result)
            print(f"✓ Report generated successfully")
            print(f"  File: {result}")
            print(f"  Size: {file_size} bytes")
            
            # Clean up
            os.remove(result)
            print("  (Test file removed)")
            return True
        else:
            print(f"✗ File not found: {result}")
            return False
            
    except Exception as e:
        print(f"✗ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_quick_generator():
    """Test QuickGenerator class."""
    print("\n" + "=" * 60)
    print("Test 7: QuickGenerator Class")
    print("=" * 60)
    
    try:
        from office_pro import QuickGenerator
        
        # Check if python-docx is available
        try:
            import docx
        except ImportError:
            print("⚠ python-docx not installed, skipping QuickGenerator test")
            return True
        
        # Create generator with custom output directory
        generator = QuickGenerator(output_dir='./test_output')
        print("✓ QuickGenerator created")
        
        # Generate a contract
        result = generator.generate(
            'contract.parking_lease',
            {
                'party_a': '王五',
                'party_b': '赵六',
                'location': 'YY大厦地下停车场',
                'space_number': 'B-123',
                'monthly_rent': 600,
                'start_date': '2024-06-01',
                'end_date': '2025-05-31'
            },
            output_filename='quick_test_contract.docx'
        )
        
        if os.path.exists(result):
            print(f"✓ Document generated with QuickGenerator")
            print(f"  File: {result}")
            
            # Clean up
            os.remove(result)
            os.rmdir('./test_output')
            print("  (Test files removed)")
            return True
        else:
            print(f"✗ File not found: {result}")
            return False
            
    except Exception as e:
        print(f"✗ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Run all tests."""
    print("\n" + "=" * 60)
    print("Office Pro Simplified API - Test Suite")
    print("=" * 60)
    
    tests = [
        ("Import Test", test_import),
        ("Dependencies Test", test_dependencies),
        ("Number Conversion Test", test_num_to_chinese),
        ("List Templates Test", test_list_templates),
        ("Generate Contract Test", test_generate_contract),
        ("Generate Report Test", test_generate_report),
        ("QuickGenerator Test", test_quick_generator),
    ]
    
    results = []
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
        except Exception as e:
            print(f"\n✗ {name} crashed: {e}")
            results.append((name, False))
    
    # Summary
    print("\n" + "=" * 60)
    print("Test Summary")
    print("=" * 60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for name, result in results:
        status = "✓ PASSED" if result else "✗ FAILED"
        print(f"  {status}: {name}")
    
    print(f"\nTotal: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n🎉 All tests passed!")
        return 0
    else:
        print(f"\n⚠ {total - passed} test(s) failed")
        return 1


if __name__ == '__main__':
    sys.exit(main())
