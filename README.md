AD User Creation - PowerShell-GUI

Here it is, my very first powershell GUI script. I understand that it's not perfect and there are improvements that could be made, however please keep in mind it is a work in progress that I will try and update when worked on. Commnents and critiquing are most welcome.

This powershell GUI was designed to help assist with onboarding new users into the orginisation, with the goal in mind to mitigate mistakes wherever possible. How I went about achieving this is to only make the admin user (Creator) type custom values into the system once, rather than having to type them in multiple times into different freeform fields meaning there was more room for error.

The Script indexes items in AD to populate dropdown boxes in the GUI (this means that the script can take a while to load depending on how much you are indexing - 1 minute), this way they again have less room for error. It even creates both an on prem and O365 mailbox (Hybrid environment) at the same time as user creation. I created two buttons in the script so you can keep it open and only have to index fields once
