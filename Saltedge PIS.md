


## a) Bank Account to Bank Account Transfers Among App Members
### User Consent and Authentication:
Each user must give consent to access their account information and initiate payments. This involves authenticating with their bank via the SaltEdge API, which usually redirects the user to their bank's login page to grant permissions securely.
### Account Selection:
Once consent is obtained, retrieve the account information of the user. The user selects which account to use for the transfer.
### Payment Initiation:
Initiate a payment order from the sender's account to the recipient's bank account details.
Specify the amount, currency, and account details (e.g., IBAN).
### Execution and Confirmation:
The API will handle the communication with the user's bank to process the transaction. 
track the payment status through SaltEdge's endpoints and confirm the transaction status to both the sender and recipient.

## b) Allowing Non-Members to Make Payments to App Members
### Payment Link Creation:
Generate a unique payment link for each transaction. 
This link should contain or encode the necessary transaction details such as the recipient's account information and the payment amount.
### Redirection to Banking App:
When a non-member clicks the payment link, they should be redirected to a SaltEdge interface (possibly a consent page) that facilitates their authentication with their own bank. 

### Payment Authorization and Initiation:
After logging in, the non-member will see the pre-filled payment details (e.g., amount, recipient’s account).
They can authorize the payment directly within their banking interface.
### Confirmation:
Once the payment is authorized and processed, you can use the SaltEdge API to confirm the transaction's completion to both the non-member (payer) and the app member (payee).

POST /api/payments
{
  "payment": {
    "creditor": {
      "account_number": "DE89370400440532013000",
      "bank_code": "DEUTDEDBBER",
      "name": "Recipient Name"
    },
    "debtor": {
      "account_number": "your customer account number"
    },
    "amount": 100.00,
    "currency": "EUR",
    "description": "Payment for services"
  }
}
