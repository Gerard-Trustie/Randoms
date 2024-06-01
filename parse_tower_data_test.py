""" This module contains tests for the parse_tower_data module. """
# in test_my_module.py
import unittest
import parse_tower_data

class TestMyModule(unittest.TestCase):
    """ Tests for the parse_tower_data module. """
    def test_function(self):
        """ Test the function."""
        #test code here
        # read the values of all the formulas in the worksheet into a dictionary
        # compare the dictionary with the expected dictionary
        # data = " test data here"  # replace with  actual test data
        formulas =[ "formula1", "formula2", "formula3"]
        column_names = ["column1", "column2", "column3"]
        result = parse_tower_data.replace_formula_references(formulas, column_names)
        expected_result = "expected result here"  # replace with the expected result
        self.assertEqual(result, expected_result)


def spare(dummy_arg=None):
    """ Spare code that can be used to test the functions."""
    # Example usage with different types of input expressions
    expressions = [
        '=IF(ANALYSYS!$C3="RTT",0,VLOOKUP(ANALYSYS!$A3,\'MW DB\'!A:AAI,151,0))',
        '=IFERROR(VLOOKUP(ANALYSYS!$A3,\'Tenant DB\'!$A:$LL,H$1,0),0)',
        '=VLOOKUP(ANALYSYS!$A3,\'Antenna DB\'!A:AW,49,0)'
    ]

    # Testing the function on each expression
    for exp in expressions:
        print(f"Input Expression: {exp}")
        print(f"Extracted VLOOKUP: {extract_vlookup_formula(exp)}\n")

    formula_example = "=VLOOKUP(ANALYSYS!$A3,'Tenant DB'!$A:$IE,I$1,0)"

    # formula_example = "=VLOOKUP(ANALYSYS!$A3,'Site DB'!$A:$O,11,0)"

    ws_name, col_name = get_lookup_sheet_and_col_name(formula_example, analysis_sheet,wb )
    print("Sheet name:", ws_name)
    print("Column name:", col_name)
    print("\n" * 3)
    return dummy_arg

if __name__ == '__main__':
    unittest.main()
