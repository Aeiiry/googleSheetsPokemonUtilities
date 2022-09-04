
function onOpen() {

  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Pokemon')
    .addItem('Get pokemon images', 'getPokemonImages')
    .addItem('Get weakness/resists', 'getPokemonTypeRelationsSheet')
    .addItem('Get pokemon type/s', 'getPokemonTypes')
    .addToUi();

}

function RegExpPokemonSpecies(string) {
  return string.match(/.*(?=-)/)[0]
}
function RegExpPokemonForm(string) {
  return string.match(/(?<=-).+/)[0]
}

function getPokemonImages() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  Logger.log('Active Range: ' + selection.getActiveRange().getA1Notation());

  var topCellLocation = [selection.getActiveRange().getRow(), selection.getActiveRange().getColumn()];

  if (activeSheet.getRange(topCellLocation[0], topCellLocation[1] + 1).getValue() != "") {
    activeSheet.insertColumnAfter(topCellLocation[1])
  }

  Logger.log('Top cell index: ' + String(topCellLocation));

  var activeList = selection.getActiveRange().getValues();
  Logger.log(activeList);
  var apiUrl = "https://pokeapi.co/api/v2/pokemon/";
  for (var i = 0; i < activeList.length; i++) {
    var pokemon = String(activeList[i]).toLowerCase();
    var response = UrlFetchApp.fetch(apiUrl + pokemon, { muteHttpExceptions: true, });
    if (response != "Not Found") {
      response = JSON.parse(response);
    }
    else {
      response = UrlFetchApp.fetch("https://pokeapi.co/api/v2/pokemon-species/" + pokemon, { muteHttpExceptions: true, })
      if (response != "Not Found") {
        response = JSON.parse(response);
        response = UrlFetchApp.fetch(apiUrl + response.varieties[0].pokemon.name, { muteHttpExceptions: true, });
      }
      else {
        var pokemonSpecies = RegExpPokemonSpecies(pokemon).toLowerCase();
        var pokemonRegion = RegExpPokemonForm(pokemon).toLowerCase();
        response = UrlFetchApp.fetch("https://pokeapi.co/api/v2/pokemon-species/" + pokemonSpecies, { muteHttpExceptions: true, });
        response = JSON.parse(response);

        var formIndex = 0;
        var found = false;
        while (found != true) {
          Logger.log("Regex result:" + response.varieties[formIndex].pokemon.name)
          if (response.varieties[formIndex].pokemon.name.match(pokemonRegion) != null) {

            Logger.log("Found regional variant name: " + apiUrl + response.varieties[formIndex].pokemon.name.match(pokemonRegion)[0]);

            response = UrlFetchApp.fetch(apiUrl + response.varieties[formIndex].pokemon.name, { muteHttpExceptions: true, });

            found = true;
          }
          else {
            formIndex++;
          }

        }

      }
      response = JSON.parse(response);
    }
    Logger.log(String(activeList[i]));
    Logger.log(response.sprites.front_default);
    Logger.log(topCellLocation[0] + 1 + i);
    Logger.log(topCellLocation[1]);


    Logger.log(activeSheet.getRange(topCellLocation[0], topCellLocation[1] + 1).getValue())

    activeSheet.getRange(topCellLocation[0] + i, topCellLocation[1] + 1).setValue("=IMAGE(\"" + response.sprites.front_default + "\")");
  }

}

function getPokemonTypes() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  Logger.log('Active Range: ' + selection.getActiveRange().getA1Notation());

  var topCellLocation = [selection.getActiveRange().getRow(), selection.getActiveRange().getColumn()];

  if (activeSheet.getRange(topCellLocation[0], topCellLocation[1] + 2).getValue() != "") {
    activeSheet.insertColumnsAfter(topCellLocation[1], 2);
  }
  else if (activeSheet.getRange(topCellLocation[0], topCellLocation[1] + 1).getValue() != "") {
    activeSheet.insertColumnsAfter(topCellLocation[1], 2);
  }


  Logger.log('Top cell index: ' + String(topCellLocation));

  var activeList = selection.getActiveRange().getValues();
  Logger.log(activeList);
  var apiUrl = "https://pokeapi.co/api/v2/pokemon/";
  for (var i = 0; i < activeList.length; i++) {
    var pokemon = String(activeList[i]).toLowerCase();
    var response = UrlFetchApp.fetch(apiUrl + pokemon, { muteHttpExceptions: true, });
    if (response != "Not Found") {
      response = JSON.parse(response);
    }
    else {
      response = UrlFetchApp.fetch("https://pokeapi.co/api/v2/pokemon-species/" + pokemon, { muteHttpExceptions: true, })
      if (response != "Not Found") {
        response = JSON.parse(response);
        response = UrlFetchApp.fetch(apiUrl + response.varieties[0].pokemon.name, { muteHttpExceptions: true, });
      }
      else {
        var pokemonSpecies = RegExpPokemonSpecies(pokemon).toLowerCase();
        var pokemonRegion = RegExpPokemonForm(pokemon).toLowerCase();
        response = UrlFetchApp.fetch("https://pokeapi.co/api/v2/pokemon-species/" + pokemonSpecies, { muteHttpExceptions: true, });
        response = JSON.parse(response);

        var formIndex = 0;
        var found = false;
        while (found != true) {
          Logger.log("Regex result:" + response.varieties[formIndex].pokemon.name)
          if (response.varieties[formIndex].pokemon.name.match(pokemonRegion) != null) {

            Logger.log("Found regional variant name: " + apiUrl + response.varieties[formIndex].pokemon.name.match(pokemonRegion)[0]);

            response = UrlFetchApp.fetch(apiUrl + response.varieties[formIndex].pokemon.name, { muteHttpExceptions: true, });

            found = true;
          }
          else {
            formIndex++;
          }

        }

      }
      response = JSON.parse(response);
    }
    Logger.log(String(activeList[i]));
    Logger.log(response.sprites.front_default);
    Logger.log(topCellLocation[0] + 1 + i);
    Logger.log(topCellLocation[1]);


    Logger.log(activeSheet.getRange(topCellLocation[0], topCellLocation[1] + 1).getValue())
    for (var t = 0; t < response.types.length; t++) {
      var type = response.types[t].type.name
      activeSheet.getRange(topCellLocation[0] + i, topCellLocation[1] + 1 + t).setValue("=IMAGE(\"" + "https://www.serebii.net/pokedex-dp/type/" + type + ".gif" + "\")");
    }
  }

  }

}

function getPokemonTypeDamageRelations(pokemon) {

  class damageRelations {

    constructor(noEffect = [], weak2x = [], weak4x = [], resist2x = [], resist4x = []) {
      this.noEffect = noEffect;
      this.weak2x = weak2x;
      this.weak4x = weak4x;
      this.resist2x = resist2x;
      this.resist4x = resist4x;
    }
  }

  var relationsObj = new damageRelations()

  var apiResponse = UrlFetchApp.fetch("https://pokeapi.co/api/v2/pokemon/" + pokemon);
  apiResponse = JSON.parse(apiResponse);

  var typesJson = apiResponse.types;
  var numTypes = typesJson.length;

  var type1DamageRelations = UrlFetchApp.fetch((typesJson[0].type.url));
  type1DamageRelations = JSON.parse(type1DamageRelations);
  type1DamageRelations = type1DamageRelations.damage_relations;

  // Get all type immunities for type 1

  if (type1DamageRelations.no_damage_from.length != 0) {
    for (var typeIndex = 0; typeIndex < type1DamageRelations.no_damage_from.length; typeIndex++) {
      relationsObj.noEffect.push(type1DamageRelations.no_damage_from[typeIndex].name)
    }
  }

  // Get all type weaknesses for type 1

  if (type1DamageRelations.double_damage_from.length != 0) {
    for (var typeIndex = 0; typeIndex < type1DamageRelations.double_damage_from.length; typeIndex++) {
      relationsObj.weak2x.push(type1DamageRelations.double_damage_from[typeIndex].name)
    }
  }

  // Get all type resistances for type 1 

  if (type1DamageRelations.half_damage_from.length != 0) {
    for (var typeIndex = 0; typeIndex < type1DamageRelations.half_damage_from.length; typeIndex++) {
      relationsObj.resist2x.push(type1DamageRelations.half_damage_from[typeIndex].name)
    }
  }

  //Return current results if pokemon is monotype

  if (apiResponse.types.length < 2) {
    return relationsObj
  }

  //Figure out combinations of weaknesses/resists if dual type
  else if (apiResponse.types.length > 1) {

    var type2DamageRelations = UrlFetchApp.fetch((typesJson[1].type.url));
    type2DamageRelations = JSON.parse(type2DamageRelations);
    type2DamageRelations = type2DamageRelations.damage_relations;

    // Get all type weaknesses for type 2 and adjust other weaknesses/resists accordingly


    if (type2DamageRelations.double_damage_from.length != 0) {

      for (var typeIndex = 0; typeIndex < type2DamageRelations.double_damage_from.length; typeIndex++) {

        //If type is found in weak2x list, remove it from that list and add it to weak4x list
        if (relationsObj.weak2x.indexOf(type2DamageRelations.double_damage_from[typeIndex].name) != -1) {

          var weakIndex = relationsObj.weak2x.indexOf(type2DamageRelations.double_damage_from[typeIndex].name);
          var typeName = relationsObj.weak2x[weakIndex];

          relationsObj.weak2x.splice(weakIndex, 1);
          relationsObj.weak4x.push(typeName);
        }
        // if weakness is in resist2x list, remove it
        else if (relationsObj.resist2x.indexOf(type2DamageRelations.double_damage_from[typeIndex].name) != -1) {
          var resistIndex = relationsObj.resist2x.indexOf(type2DamageRelations.double_damage_from[typeIndex].name);
          relationsObj.resist2x.splice(resistIndex, 1);
        }
        else {
          relationsObj.weak2x.push(type2DamageRelations.double_damage_from[typeIndex].name);
        }
      }
    }

    // Get all type resistances for type 2 and adjust other weaknesses/resists accordingly

    if (type2DamageRelations.half_damage_from.length != 0) {

      for (var typeIndex = 0; typeIndex < type2DamageRelations.half_damage_from.length; typeIndex++) {

        var currentResist = type2DamageRelations.half_damage_from[typeIndex].name;

        //if current resist is in weak2x list, remove it
        if (relationsObj.weak2x.indexOf(currentResist) != -1) {
          var weakIndex = relationsObj.weak2x.indexOf(currentResist)
          relationsObj.weak2x.splice(weakIndex, 1);
        }
        //if current resist is in weak4x list, move it to weak2x list
        else if (relationsObj.weak4x.indexOf(currentResist) != -1) {
          var weakIndex = relationsObj.weak4x.indexOf(currentResist);
          relationsObj.weak4x.splice(weakIndex, 1);
          relationsObj.weak2x.push(currentResist);
        }
        //if current resist is in resist2x list, move it to resist4x list
        else if (relationsObj.resist2x.indexOf(currentResist) != -1) {
          var resistIndex = relationsObj.resist2x.indexOf(currentResist);
          relationsObj.resist2x.splice(resistIndex, 1);
          relationsObj.resist4x.push(currentResist);
        }
        else {
          relationsObj.resist2x.push(currentResist);
        }
      }
      // Get all type immunities for type 2, remove any of those types from other weak/strong lists

      if (type2DamageRelations.no_damage_from.length != 0) {
        for (var typeIndex = 0; typeIndex < type2DamageRelations.no_damage_from.length; typeIndex++) {
          var noEffectName = type2DamageRelations.no_damage_from[typeIndex].name;

          if (relationsObj.resist2x.indexOf(noEffectName) != -1) {
            relationsObj.resist2x.splice(relationsObj.resist2x.indexOf(noEffectName), 1)
          }
          else if (relationsObj.resist4x.indexOf(noEffectName) != -1) {
            relationsObj.resist4x.splice(relationsObj.resist4x.indexOf(noEffectName), 1)
          }
          else if (relationsObj.weak2x.indexOf(noEffectName) != -1) {
            relationsObj.weak2x.splice(relationsObj.weak2x.indexOf(noEffectName), 1)
          }
          else if (relationsObj.weak4x.indexOf(noEffectName) != -1) {
            relationsObj.weak4x.splice(relationsObj.weak4x.indexOf(noEffectName), 1)
          }
          relationsObj.noEffect.push(noEffectName);
        }
      }
    }
  }
  // double check first type immunities
  if (type1DamageRelations.no_damage_from.length != 0) {
    for (var typeIndex = 0; typeIndex < type1DamageRelations.no_damage_from.length; typeIndex++) {
      var noEffectName = type1DamageRelations.no_damage_from[typeIndex].name;
      if (relationsObj.resist2x.indexOf(noEffectName) != -1) {
        relationsObj.resist2x.splice(relationsObj.resist2x.indexOf(noEffectName), 1);
      }
      else if (relationsObj.resist4x.indexOf(noEffectName) != -1) {
        relationsObj.resist4x.splice(relationsObj.resist4x.indexOf(noEffectName), 1);
      }
      else if (relationsObj.weak2x.indexOf(noEffectName) != -1) {
        relationsObj.weak2x.splice(relationsObj.weak2x.indexOf(noEffectName), 1);
      }
      else if (relationsObj.weak4x.indexOf(noEffectName) != -1) {
        relationsObj.weak4x.splice(relationsObj.weak4x.indexOf(noEffectName), 1);
      }
    }
  }
  return relationsObj;
}

function getPokemonTypeRelationsSheet() {

  var global = {};

  var relationsList = ["weak2x", "weak4x", "resist2x", "resist4x", "noEffect"];
  var relationsListDescStr = ["2x weak to", "4x weak to", "2x resistant to", "4x resistant to", "not affected by"]

  var typeList = ["normal", "fire", "water", "grass", "flying", "fighting", "poison", "electric", "ground", "rock", "psychic", "ice", "bug", "ghost", "steel", "dragon", "dark", "fairy"]

  for (var relation in relationsList) {
    global[relationsList[relation] + "Count"] = {}
    for (var type in typeList) {
      global[relationsList[relation] + "Count"][typeList[type]] = []
    }

  }


  var activeSheet = SpreadsheetApp.getActiveSheet();
  var selection = activeSheet.getSelection();
  Logger.log('Active Range: ' + selection.getActiveRange().getA1Notation());

  var topCellLocation = [selection.getActiveRange().getRow(), selection.getActiveRange().getColumn()];

  if (activeSheet.getRange(topCellLocation[0], topCellLocation[1] + 1).getValue() != "") {
    activeSheet.insertColumnAfter(topCellLocation[1])
  }

  Logger.log('Top cell index: ' + String(topCellLocation));

  var activeList = selection.getActiveRange().getValues();
  Logger.log(activeList);
  var apiUrl = "https://pokeapi.co/api/v2/pokemon/";
  for (var i = 0; i < activeList.length; i++) {
    var pokemon = String(activeList[i]).toLowerCase();
    var response = UrlFetchApp.fetch(apiUrl + pokemon, { muteHttpExceptions: true, });
    if (response != "Not Found") {
      response = JSON.parse(response);
    }
    else {
      response = UrlFetchApp.fetch("https://pokeapi.co/api/v2/pokemon-species/" + pokemon, { muteHttpExceptions: true, })
      if (response != "Not Found") {
        response = JSON.parse(response);
        response = UrlFetchApp.fetch(apiUrl + response.varieties[0].pokemon.name, { muteHttpExceptions: true, });
      }
      else {
        var pokemonSpecies = RegExpPokemonSpecies(pokemon).toLowerCase();
        var pokemonRegion = RegExpPokemonForm(pokemon).toLowerCase();
        response = UrlFetchApp.fetch("https://pokeapi.co/api/v2/pokemon-species/" + pokemonSpecies, { muteHttpExceptions: true, });
        response = JSON.parse(response);

        var formIndex = 0;
        var found = false;
        while (found != true) {
          Logger.log("Regex result:" + response.varieties[formIndex].pokemon.name)
          if (response.varieties[formIndex].pokemon.name.match(pokemonRegion) != null) {

            Logger.log("Found regional variant name: " + apiUrl + response.varieties[formIndex].pokemon.name.match(pokemonRegion)[0]);

            response = UrlFetchApp.fetch(apiUrl + response.varieties[formIndex].pokemon.name, { muteHttpExceptions: true, });

            found = true;
          }
          else {
            formIndex++;
          }

        }

      }
      response = JSON.parse(response);
    }
    var damageRelations = getPokemonTypeDamageRelations(response.name)
    var damageRelationsStr = ""
    if (damageRelations.resist2x.length > 0) {
      damageRelationsStr += "Resists x 2:"

      for (var s = 0; s < damageRelations.resist2x.length; s++) {
        damageRelationsStr += " " + damageRelations.resist2x[s]
      }
    }
    if (damageRelations.resist4x.length > 0) {
      damageRelationsStr += "\nResists x 4:"

      for (var s = 0; s < damageRelations.resist4x.length; s++) {
        damageRelationsStr += " " + damageRelations.resist4x[s]
      }
    }

    if (damageRelations.weak2x.length > 0) {
      damageRelationsStr += "\nWeak x 2:"

      for (var s = 0; s < damageRelations.weak2x.length; s++) {
        damageRelationsStr += " " + damageRelations.weak2x[s]
      }
    }

    if (damageRelations.weak4x.length > 0) {
      damageRelationsStr += "\nWeak x 4:"

      for (var s = 0; s < damageRelations.weak4x.length; s++) {
        damageRelationsStr += " " + damageRelations.weak4x[s]
      }
    }
    if (damageRelations.noEffect.length > 0) {
      damageRelationsStr += "\nNot affected by:"

      for (var s = 0; s < damageRelations.noEffect.length; s++) {
        damageRelationsStr += " " + damageRelations.noEffect[s]
      }
    }

    activeSheet.getRange(topCellLocation[0] + i, topCellLocation[1] + 1).setValue(damageRelationsStr);

    Logger.log(damageRelations)

    for (relation in relationsList) {

      for (type in damageRelations[relationsList[relation]]) {

        global[relationsList[relation] + "Count"][damageRelations[relationsList[relation]][type]].push(pokemon)

      }
    }

  }


  for (relation in relationsList) {
    global[relationsList[relation] + "CountDescStr"] = ""
    var currentRelationCount = relationsList[relation] + "Count"

    global[currentRelationCount] = Object.entries(global[currentRelationCount])
      .sort(
        (a, b) => a[1].length - b[1].length
      )
      .reverse()
      .reduce((r, [k, v]) => ({ ...r, [k]: v }), {});

    if (Object.entries(global[currentRelationCount])[0][1].length != 0) {

      for (type in global[currentRelationCount]) {

        if (global[currentRelationCount][type].length == 1) {

          global[relationsList[relation] + "CountDescStr"] += global[currentRelationCount][type].length + " (" + String(Math.round((global[currentRelationCount][type].length / activeList.length) * 100) * 100) / 100 + "%) Pokemon is " + relationsListDescStr[relation] + " " + type + "\n"
        }
        else if (global[currentRelationCount][type].length > 1){

          global[relationsList[relation] + "CountDescStr"] += global[currentRelationCount][type].length + " (" + String(Math.round((global[currentRelationCount][type].length / activeList.length) * 100) * 100) / 100 + "%) Pokemon are " + relationsListDescStr[relation] + " " + type + "\n"
        }
      }
      Logger.log(global[relationsList[relation] + "CountDescStr"])




      activeSheet.getRange(topCellLocation[0]+activeList.length,topCellLocation[1]+parseInt(relation))
      .setValue(global[relationsList[relation] + "CountDescStr"])

    }

  }

}

















