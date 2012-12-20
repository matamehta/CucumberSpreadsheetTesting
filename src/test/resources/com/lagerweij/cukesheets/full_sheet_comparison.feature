Feature: Full sheet comparison

Scenario: Matching sheets
    Given the input sheet input_sheet_1.xls
    When the file is processed
    Then the output is exactly the same as baseline_sheet_1_matching.xls

Scenario: Non-matching sheets
    Given the input sheet input_sheet_1.xls
    When the file is processed
    Then the output is not the same as baseline_sheet_1_not_matching.xls

Scenario: Failure with non-matching sheets
    Given the input sheet input_sheet_1.xls
    When the file is processed
    Then the output is exactly the same as baseline_sheet_1_not_matching.xls
