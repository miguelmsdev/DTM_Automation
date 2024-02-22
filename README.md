# DTM_Automation
This project automates the task of manually inputting info on the Debris Transport Manifests (In Portuguese, MTR - Manifesto de Transporte de Resíduos, but we'll call it DTM), a licensed government system made for monitoring and keeping tabs on Senders, Transporters, and Receivers involved in construction produced materials.

The authorizations come in an A4 paper, which is scanned with a barcode scanner for its identification number, after having an Excel sheet with said authorizations you just run the program, you can let it deal with the receiving date so it puts the same date the authorization was issued, or you can modify the receiving date yourself.

I've had up to a thousand papers coming a day, and it would usually take me about 11 seconds per DTM, this program does it in 1 second, which is 90% faster, turning what could be 3 hours of repetitive work into 10 min of monitoring the program doing it, that way I saved the company cost and time it would waste by assigning extra paid workforce.

The project was written using Python, Selenium, and some small generic libraries,  took me around 5 days working part-time to build the logic and write the code, the same time it would take the code to pay for itself time-wise, and from there on to bring profit for the company and more free time for me to invent new solutions while being available for more brain-needed challenges.

It is far from full-fledged and needs constant maintenance as does every third-party dependant software, the website is constantly changing references and elements' paths, sometimes refactoring almost all of it.
I've paused updating the code after a year and a half of use since contracts have hit the estimated deadline, but it proved useful saving time and I learned a lot doing it.
