Feature: Checking a specific value

Scenario: Check Pi
    Given the input sheet input_sheet_1.xls
    When the file is processed
    Then the output is generated in baseline_sheet_1_matching.xls
    And cell at coordinate 2, 3 has numeric value 3.14159265359

Scenario Outline: Check numeric values
    Given the input sheet <input_sheet>
    When the file is processed
    Then the output is generated in <output_sheet>
    And cell at coordinate <column>, <row> has numeric value <numeric_value>

Examples:
    | input_sheet       | output_sheet                      | column    | row   | numeric_value |
    | input_sheet_1.xls | baseline_sheet_1_matching.xls     | 2         | 3     | 3.147123      |
    | input_sheet_1.xls | baseline_sheet_1_not_matching.xls | 2         | 3     | 3.14159265359 |


Scenario Outline: Check string values
    Given the input sheet <input_sheet>
    When the file is processed
    Then the output is generated in <output_sheet>
    And cell at coordinate <column>, <row> has string value <string_value>

Examples:
    | input_sheet       | output_sheet                      | column    | row   | string_value |
    | input_sheet_1.xls | baseline_sheet_1_not_matching.xls | 0         | 4     | row 3         |
