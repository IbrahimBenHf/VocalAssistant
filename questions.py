questions_en = {
    "PSD in creation, what do you want as a title?": 1,
    "what is the client context?": 1,
    "what is the business context?": 1,
    "give me a brief description of the change request ?": 1,
    "do you want to add new feature description ?": 1,
    "what is the feature's title ?": 1,
    "what is the description for this feature ?": 1,
    "PSD is now saved and is gonna be sent to you on mail.": 0,

    "PFR in creation, what do you want as a title?": 2,
    "what is the aim of the document?": 2,
    "describe the current behavior?": 2,
    "what is the proposed solution?": 2,
    "do you want to add another new feature ?": 2,
    "what is the new feature's title ?": 2,
    "what is the description ?": 2,
    "PFR is now saved and is gonna be sent to you on mail.": 0,

    "test plan in creation, what is the title of the test?": 3,
    "what is the description of the test case?": 3,
    "what is the actual status of this test case ?": 3,
    "do you want to add another test case ?": 3,
    "test plan is now saved and is gonna be sent to you on mail.": 0,

}


def get_id(question):
    if question in questions_en:
        return questions_en[question]
    else:
        return 0

