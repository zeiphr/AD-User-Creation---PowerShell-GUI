AD User Creation - PowerShell GUI

Introducing my first PowerShell GUI script designed to streamline the process of onboarding new users within the organization. While I acknowledge that there is room for improvement, please consider that this script is a work in progress, and I will continue to refine it. I welcome your comments and critiques.

The primary objective of this PowerShell GUI is to minimize the potential for errors during the user creation process. To achieve this, I have implemented a system where the admin user (Creator) only needs to input custom values once. This eliminates the need to repetitively enter information into various freeform fields, reducing the likelihood of mistakes.

The script leverages Active Directory (AD) indexing to populate dropdown boxes within the GUI. This indexing process may take some time to complete, depending on the volume of data being indexed, typically around 1 minute. By doing so, we further mitigate the risk of errors in user data entry.

Additionally, the script has the capability to create both on-premises and Office 365 mailboxes, making it especially valuable in a hybrid environment. To enhance usability, I have incorporated two buttons into the script, allowing users to keep it open and index fields just once, improving efficiency in user provisioning.
