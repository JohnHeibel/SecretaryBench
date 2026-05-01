
Hello everybody. Well done in creating the base implementation of the system. These next steps are a lot more iterative and broad, and the task is due on Tuesday 5th by the end of the day.

There are two pieces that are currently separate from each other:
`Email parsing Pipeline` and `Model instruction set-up & tooling`

The main task to complete is bringing these two pieces together. Currently we have the `model_interaction_mock` function (line 160) acting as the model processor. This should be replaced by using Miguel's system instead.

This is what the system should ultimately do:

1) Read emails
2) Parse emails
3) Serve/feed emails to the model
4) Emails get moved to pools (served email pool, inactive email pool, active scenario pool, etc…) 
5) This is with the purpose of keeping track of which scenarios are complete.
6) Model tool calls
7) Model inputs data into Calendar/Todo Python objects
8) Objects are graded for every iteration of emails “served”
9) Simulate 100 days

What still needs to get done:
- An email served from a specific scenario cannot be served unless previous emails in that chain have already been served.
-> For example: Emails ABCDE are in a single scenario ID. Then email D cannot be served unless ABC emails have been already served.
- Emails-to-model pipeline
- A system prompt of instructions for the AI to know what it’s expectations are
- Grading system with real emails fed
- No leftover emails after 100 days

We will have a meeting on Thursday to test the system. It should be fully functional with basic emails.

This task is due Tuesday 5th.

Please ask me (Daniel) any questions.
