# googleSheetsPokemonUtilities

A fairly simple (right now) utility for inserting data about pokemon in to google sheets via apps script and pokeapi (https://pokeapi.co/)

Currently has 3 functions, all require an input of one or more pokemon names selected.

To use any of the functions simply select the pokemon you want (they must be in a column) and select the function you want in the "Pokemon" menu in the toolbar

The scripts will attempt to find the closest matching pokemon name in the pokeapi database, as there are some unintuitive naming conventions such as pokemon with multiple base forms needing to have the form specified (wormadam is invalid, wormadam-trash is valid). If you're unsure just enter the name of the pokemon and it will return its first form

If you can't seem to get it to recognize a pokemon, head to https://pokeapi.co/ and enter pokemon-species/ and then the name of the pokemon's species (e.g wormadam), then check under "varieties" for a list of pokemon names the api will recognise. 

If some of the names are odd or frustrating to enter, message @Aeiiry and I'll try to improve the input checking
