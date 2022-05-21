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

    "What should I write?":4,

    "what do you want to translate?":5,

    "what is the to do to add?":6,

    "what is the number of to do to complete?":7

}


questions_fr = {
    "PSD en création, tu veux quoi comme titre ?": 1,
    "quel est le contexte client ?": 1,
    "quel est le contexte commercial ?": 1,
    "me donner une brève description de la demande de changement ?": 1,
    "voulez-vous ajouter une nouvelle fonctionnalité?": 1,
    "quel est le titre de la fonctionnalité ?": 1,
    "quelle est la description de cette fonctionnalité ?": 1,
    "PSD est maintenant enregistré et va vous être envoyé par courrier.": 0,

    "PFR en création, tu veux quoi comme titre ?": 2,
    "quel est l'objectif du document ?": 2,
    "décrire l'existant ?": 2,
    "quelle est la solution proposée ?": 2,
    "voulez-vous ajouter une autre nouvelle fonctionnalité ?": 2,
    "quel est le titre de la nouvelle fonctionnalité?": 2,
    "c'est quoi le descriptif ?": 2,
    "PFR est maintenant enregistré et va vous être envoyé par courrier.": 0,

    "plan de test en création, quel est le titre du test ?": 3,
    "quelle est la description du cas de test ?": 3,
    "quel est le statut actuel de ce cas de test ?": 3,
    "voulez-vous ajouter un autre cas de test ?": 3,
    "Le plan de test est maintenant enregistré et va vous être envoyé par courrier.": 0,

    "Que devrais-je écrire?":4,

    "tu veux traduire quoi ?":5,

    "quelle est la tâche à ajouter?":6,

    "quel est le numéro de la tâche à accomplir?":7

}





def get_id(question):
    if question in questions_en:
        return questions_en[question]
    else:
        return 0


def get_id_fr(question):
    if question in questions_fr:
        return questions_fr[question]
    else:
        return 0

