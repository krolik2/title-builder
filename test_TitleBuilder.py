import pytest
from TitleBuilder import BuildTitle


@pytest.fixture
def build_title_instance():
    return BuildTitle(
        asin="ASIN123",
        brand="Brand",
        department="Department",
        model_name="Model_Name",
        model_num="Model_Number",
        color="Red",
        size="Large",
        flavor="Tasty",
        material="Metal",
        part_num="Part_Number",
        wattage="100",
        voltage="220",
        item_name="Item_Name",
        cpu_model="CPU Model",
        computer_memory="8GB",
        memory_storage_capacity="256GB",
        hard_disk="SSD",
        graphics_description="Integrated Graphics",
        operating_system="Windows 10",
        keyboard_layout="QWERTY",
        sub_brand="Sub Brand",
    )

# Test the build_title_group1 method
def test_build_title_group1(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Department Model_Name Item_Name, Red, Large",
    }
    assert build_title_instance.build_title_group1() == expected_output

# Test the build_title_group2 method
def test_build_title_group2(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Department Model_Name Model_Number Item_Name, Red, Large",
    }
    assert build_title_instance.build_title_group2() == expected_output

# Test the build_title_group3 method
def test_build_title_group3(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Name Sub Brand Item_Name, Red, Large",
    }
    assert build_title_instance.build_title_group3() == expected_output

# Test the build_title_group5 method
def test_build_title_group5(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Name Item_Name, Red, Large",
    }
    assert build_title_instance.build_title_group5() == expected_output

# Test the build_title_group6 method
def test_build_title_group6(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Name Item_Name, Red, Large, 100W 220V",
    }
    assert build_title_instance.build_title_group6() == expected_output

# Test the build_title_group7 method
def test_build_title_group7(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Name Item_Name, Red, 100W 220V",
    }
    assert build_title_instance.build_title_group7() == expected_output

# Test the build_title_group8 method
def test_build_title_group8(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Name Item_Name, Tasty, Large",
    }
    assert build_title_instance.build_title_group8() == expected_output

# Test the build_title_group9 method
def test_build_title_group9(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Name Item_Name, Metal, Red, Large",
    }
    assert build_title_instance.build_title_group9() == expected_output

# Test the build_title_group10 method
def test_build_title_group10(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Name Item_Name, Large",
    }
    assert build_title_instance.build_title_group10() == expected_output

# Test the build_title_group11 method
def test_build_title_group11(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Name Model_Number Item_Name, Red, Large",
    }
    assert build_title_instance.build_title_group11() == expected_output

# Test the build_title_group12 method
def test_build_title_group12(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Name Model_Number Item_Name, CPU Model, 8GB RAM, 256GB SSD, Integrated Graphics, Windows 10, QWERTY, Red, Large"
    }
    assert build_title_instance.build_title_group12() == expected_output

# Test the build_title_group14 method
def test_build_title_group14(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Model_Number Item_Name, Red, Large",
    }
    assert build_title_instance.build_title_group14() == expected_output

# Test the build_title_group15 method
def test_build_title_group15(build_title_instance):
    expected_output = {
        "asin": "ASIN123",
        "title": "Brand Part_Number Item_Name, Red, Large",
    }
    assert build_title_instance.build_title_group15() == expected_output

