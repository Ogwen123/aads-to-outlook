# aads-to-outlook
A script to automatically add your AA Driving School lessons to Outlook.
It will not add past, cancelled or duplicate lessons (unless you mess with done.json)

## data needed
|field|explanation|
|-----|-----------|
|aads.auth | the authorization header in the 'da-api.theaa.digital' graphql request, you will need to update this regularly|
|aads.id | can be found in the 'da-api.theaa.digital' graphql request payload under variables.learnerId|
|outlook.calendar_id | under Body.SavedItemFolder.BaseFolderId.Id in the CreateCalendarEvent request|
|outlook.cookies | the entire cookies header in the CreateCalendarEvent request|
|outlook.email | your email|
