# Overview

This project adds support to Outlook to automatically delay the delivery of emails that are sent outside of business hours.

# Installation

Download and run the installer for the current release. Note, this is not signed. Alternatively, you can build from source and publish your own installer.

To confirm that the add-in is installed, navigate to **File**, **Options**, and select **Add-ins**. You can then click the **Go...** button to manage COM Add-ins and confirm that the add-in is installed and enabled as shown here:

![Add-in enablement status](doc\addin.jpg)

# Configuration

This add-in supports customizing the dates and times that are outside of business hours. This configuration should roam with your user profile so that you don't need to redefine it on every machine where you use Outlook. You can do this by accessing the **Add-ins** tab of the Outlook window as shown here:

![Configuration via explorer ribbon](doc\explorer_ribbon.jpg)

In the event that you need to override delay delivery for an email being sent outside of business hours, you can do so via the **Add-ins** tab for a mail item as shown here:

![Overriding delay send via mail ribbon](doc\mail_ribbon.jpg)

# How it works

This feature is implemented as an [Outlook VSTO Add-in](https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-outlook?view=vs-2019). It adds a custom handler for when mail items are being sent and checks to see if the mail is being sent outside of the sender's defined business hours. If so, it sets a deferred delivery date for the email to align with the start of the next business day (which may be after a weekend).

# FAQ

1. Does Outlook need to remain open for delayed mails to be sent? Yes.

2. Is it possible to delay delivery to align with receipent work schedules? Not with this add-in. For this to work, it would likely require integration with Exchange.

# References

Others have identified the need for this Outlook feature and have written tools to help. This blog from 2017 showed how [delay send could be implemented using VBA](https://medium.com/@BMatB/delaying-email-sending-outlook-vba-dbfd41a6ad01).