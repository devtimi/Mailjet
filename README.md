# MailJet Class for Xojo
Class for sending transactional email via MailJet in pure Xojo Code. This MailJet class sends `EmailMessage` objects in a similar manner to `SMTPSecureSocket`.

## Usage

1. Set up your MailJet account and [collect your API credentials](https://app.mailjet.com/account/api_keys). These credentials grant full access to your account, from sending actual email to statistics and information. Do not store your API credentials in a binary destined for distribution.
2. Copy `MailJet` and `MailJetException` into your project.
3. Configure the MailJet class constants with your API credentials.
4. Craft an `EmailMessage` object that represents the email you wish to send.
5. Add the message to the `Messages` array on an instance of the `MailJet` class.
6. Send mail!

**Demo App Example**

The demo app window shows a simple example of sending an email message. This class is designed for single transactional emails. If you append multiple recipients they will all receive the same email, rather than individualized emails (like for a newsletter). The MailJet service offers campaign features, but this class is not an implementation of these features.

**A note about Name + Email**

MailJet allows developers to send email with the recipient's name to improve deliverability. However, Xojo `EmailMessage` does not have a separate name and email field. To work around this limitation, this class parses addresses using the normal email convention of `Person Name <address@domain.com>`

## Events

#### Error
The `Error` event will be raised for each and every error that occurs. This is because the class can continue attempting to send in most cases.

#### Mail Sent
If at least one email was successfully sent, the `MailSent` event will be raised upon completion.