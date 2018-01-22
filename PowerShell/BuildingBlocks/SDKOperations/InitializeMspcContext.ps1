# This script illustrates how to use the Initialize-MSPC_Context cmdlet.

# Case 1: Initialize an mspc context with a customer ID
# Note: the customer's workgroup is used in creating both the workgroup and workgroup ticket.
Initialize-MSPC_Context -Credentials $credentials -CustomerId 'your customer ID here'

# Case 2: Initialize an mspc context with a workgroup ID
# Note: no customer nor customer ticket are created in this case.
Initialize-MSPC_Context -Credentials $credentials -WorkgroupId 'your workgroup ID here'

# Case 3: Clear the existing global $mspc context
# Note: clears the existing $mspc context before creating a new mspc context.
Initialize-MSPC_Context -Clear

(INNER JOIN MailboxConnectors SecurityMailboxConnectors_Mailbox WITH (NOLOCK) ON Mailbox.ConnectorId = SecurityMailboxConnectors_Mailbox.ConnectorId AND (SecurityMailboxConnectors_Mailbox.IsDeleted <> @SecurityConnectorIsDeleted0))