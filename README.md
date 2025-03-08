# Coding Challenge

## Description

This is not as much a coding challenge as it is a coding exercise.

The code in this repository is a small snippet of production code that at one point
was used in a real-world project. It works and has no intentional bugs, but it's
not perfect. Your task is to refactor the code to make it **better**.

"Better" is intentionally vague. You can interpret it in any way you like. You can
make the code more readable, more efficient, more maintainable, or anything else
you think is important. You can also add tests if you think that would be helpful.

Only the `DocumentProcessingService` class is relevant to this exercise, other
classes are there to provide context and mock dependencies.

## Instructions

1. Clone this repository.
2. Refactor the code in the `DocumentProcessingService` class (and any other classes
   if you think it's necessary).
3. Commit your changes to a new branch.
4. Push the branch to the repository.
5. Create a pull request.

## Notes

- You can use any tools or libraries you like. But if possible, make notes of the
  tools you used and why you chose them.

## Refactor Notes

 - Refactored each methods in DocumentProcessingService
 - Injected useful classes and libraries to do it more readeble and testable
 - Splitted smaller methods, removed unnecesserary objects ..
 - Could be still need to do more refactor, that's why I dropped small inline comments where it could be neccessery to improve
 - Used moq library to mock some data or if it'd be necessary for future developments.
 - Added more unit tests and still could be needed to add new unit test cases.(Also dropped comments for assertions and UT methods but could not complete since the time concerns)

**Good luck!**
