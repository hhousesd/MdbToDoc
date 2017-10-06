SELECT
    [Case].GWIcaseRef_ID AS CaseId, Case.strCase_Name AS CaseName,
    Role.strCaseRelation AS Role,
    Person.strFirstName_client AS FirstName, Person.strLastName_client AS LastName,
    Person.strPhoneNumber, Person.strSex as Sex,
    Person.strAddress as AddressStreet, Person.strCity as AddressCity, Person.strState as AddressState,
        Person.strPostalCode as AddressPostalCode, Person.strCountry as AddressCountry,
    Person.memNotes as Notes
FROM ((tblPhonelog AS Log
INNER JOIN tblRSVPCases AS [Case] ON Log.GwiCaseRef_ID = Case.GwiCaseRef_ID)
INNER JOIN tblPeople AS Person ON Log.GwiCaseRef_ID = Person.GWIcaseRef_ID)
INNER JOIN tblCaseRelation AS Role ON Person.CaseRelationID = Role.CaseRelationID
