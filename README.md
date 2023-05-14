# YGOPRO-pack-bot

This is an automation bot, made with python, using playwright and Excel for YGOPRO pack simulator. 

This bot will login into YGOPRO (just fill the specific fields with your email + password), filter the pack you want to open, open all the packs and insert the cards into Excel. This process is repeated for 11 times
Later, the bot will wait 50 minutes for you to analyze all the 11 packs on Excel, and if you want, you can add to your collection or discard the pack opened.


Requirements:
- Profile on YGOPRO
- Openpyxl, to deal with Excel
- Playwright, for opening the internet browser
