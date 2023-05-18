# Engage Demo - Validation

## Form

1. Create a Form called "Person".
1. Add a text field called "FirstName".
1. Add a text field called "LastName".
1. Add a number field called "Age".
1. Add a radio button field called "Gender" and choose valid options.

## Globals

1. In the Globals area, add `Option Declare`.
1. Add `Use "VoltScript Testing".

## QuerySave

VoltScriptTesting can not only be used for unit / integration testing, it can also be used for validation.

If we suppress the report, we can write unit tests to test the field values. If the tests ran successfully, the document is valid and we can continue saving. If not, we abort saving and return the test failures.

1. Declare a variable `doc` as a NotesDocument.
1. Set it to `Source.Document`.
1. Declare a string variable called `errors`.
1. Declare a boolean variable called `validateAge`.
1. Add error handling to report any error and exit the sub.
1. Declare a `testSuite` variable as a new TestSuite, giving it the name "Validate Doc". The name will not be used, but is required.
1. Add a test to validate that `doc.FirstName(0)` is more than one character. `assertTrue` can be used to test the length, using `FullTrim()` and `Len()`.
1. Add a test to validate that `doc.LastName(0)` is not an empty string.
1. Set the boolean `validateAge` to the result of a test called "Check age is completed", validating the `CStr(doc.Age(0))` is not an empty String.
1. If `validateAge` is true, perform three more tests.
    1. Add a test to validate `doc.Age(0)` is greater than 0.
    1. Add a test to validate `doc.Age(0)` is less than 110.
    1. Add a test to validate that `doc.Age(0)` is a whole number. `Fraction()` is a LotusScript function that returns the fractional portion of a number. (For a whole number, the fractional portion should be 0.)
1. Add a test to validate `doc.Gender(0)` is not an empty string.
1. Set `Continue` to the result of `testSuite.ranSuccessfully()`.
1. If `continue` is False, add code to loop through the results and capture the descriptions for any tests that did not pass.
    1. Add a `ForAll` loop to iterate over `testSuite.results`.
    1. For each test, check if the result is not "Passed". If so, Set `errors` to its existing value, the test description, and a new line.
    1. After the ForAll loop, `MsgBox` the errors to the user.

!!! success
    You have successfully added validation for the Form. If you add fields or change fields in the future, you just need to add additional tests. No other changes are needed.

!!! tip
    If you have subforms and want to validate those separately, it may be preferable to move the QuerySave code to a script library. Alternatively, you could set the TestSuite as a global variable.

[Full Code](../assets/example_code/querySave.lss)