# O365BulkShare
I wrote this script as a means for bulk setup of existing vendors for our extranet on SharePoint Online. One of the biggest issues with the current flow for adding external users is that the invitation emails get sent from Microsoft and the content is mostly static. In my experience, many external users may disregard the email as junk mail or their email servers may even flag it automatically as such. External users expect to see emails from your domain, not Microsoft's. This script will send an email via your Exchange Online with the content of your choosing (including HTML.) The CSV file should contain just one column of data, external users' email addresses, named Email.