# googleSheetsPokemonUtilities

A fairly simple (right now) utility for inserting data about pokemon in to google sheets via apps script and pokeapi (https://pokeapi.co/)

## Installation

**NOTE:** I'm a stranger on the internet, if you don't want to give the script access to your sheet/account just don't install it. All the code can be viewed  [here](https://github.com/Aeiiry/googleSheetsPokemonUtilities/blob/main/code.gs) in the code.gs file. All the script does is call pokeapi and modify some cells in your google sheet.

1. If you haven't already, make a new google sheet [here](https://sheets.google.com/)
2. In the toolbar in your sheet, click "Extensions" and then click "Apps Script" - This will open a new window
3. Copy the contents of the [code.gs file](https://github.com/Aeiiry/googleSheetsPokemonUtilities/blob/main/code.gs) to code.gs within the apps script page, making sure to copy it below any existing text
4. Press control+S or click the save icon, go back to your google sheet
5. Wait for a little bit and refresh the page, a button should show up near the top of the page labelled "Pokemon". If it doesn't show up either refresh again and/or wait for ~20 seconds
6. Next we need to grant the script access to your google sheet. Click "Pokemon" in the toolbar and then click any of the buttons that pop up
7. A window will pop up reading "Authorization Required", click "Continue". A new window will pop up
8. Select your google account, click "Advanced", click "Go to [your project name] (unsafe)", then click "Allow"
9. Now the scripts will work!

## Scripts

Currently has 3 functions, all require an input of one or more pokemon names selected om a column (just by clicking and dragging in google sheets).

### Get pokemon Images
Will insert images of the given pokemon in the cell to the right. Inserts a column if there is not enough room

### Get pokemon type/s
Inserts 1-2 images of the given pokemon's types in the column/s to the right of it. Inserts two columns if there is not enough room.

### Get pokemon weakness/resists - WARNING WORK IN PROGRESS/BAD
Gets a tally of each pokemon's weaknesses and resists, additionally inserts a summary of all selected pokemon's weaknesses/resists to see common types for each.

###

To use any of the functions simply select the pokemon you want (they must be in a column) and select the function you want in the "Pokemon" menu in the toolbar

The scripts will attempt to find the closest matching pokemon name in the pokeapi database, as there are some unintuitive naming conventions such as pokemon with multiple base forms needing to have the form specified (wormadam is invalid, wormadam-trash is valid). If you're unsure just enter the name of the pokemon and it will return its first form

If you can't seem to get it to recognize a pokemon, head to https://pokeapi.co/ and enter pokemon-species/ and then the name of the pokemon's species (e.g wormadam), then check under "varieties" for a list of pokemon names the api will recognise. 

If some of the names are odd or frustrating to enter, create an issue and I'll try to improve the input checking
