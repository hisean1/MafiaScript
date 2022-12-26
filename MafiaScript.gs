const NAME_ROW = 5;
const ROLE_ROW = 6;
const VISIT_ROW = 7;
const ALIVE_ROW = 8;
const PRIORITY_ROW = 9;
const USES_ROW = 10;
const MAX_PLAYERS = 10;
const ERROR_CELL_OFFSET = 1;
const ROW_TITLE_OFFSET = 2;
const PLAYER_INDEX_OFFSET = 3;
const ROLE_ORDER_OFFSET = 13;
const GAMEPLAY_INFO_OFFSET = 15;
const NUM_ROLES = 14; // Treats non-targeting roles as one, as well as all mafia support

const ss = SpreadsheetApp.getActive().getSheetByName('Mafia');
const NUM_PLAYERS_CELL = ss.getRange(NAME_ROW - 1, ROW_TITLE_OFFSET + 1);
const ROLE_ERROR_CELL = ss.getRange(ROLE_ROW, ERROR_CELL_OFFSET);
const VISIT_ERROR_CELL = ss.getRange(VISIT_ROW, ERROR_CELL_OFFSET);
const NIGHT_ERROR_CELL = ss.getRange(1, ERROR_CELL_OFFSET);
const CURRENT_PLAYER_CELL = ss.getRange(6, GAMEPLAY_INFO_OFFSET);

function prepareGame()
{
  clearLastGame();

  // Assign priorities
  for (var i = 0; i < NUM_PLAYERS_CELL.getValue(); i++)
  {
    revivePlayer(i + PLAYER_INDEX_OFFSET);

    var priocell = ss.getRange(PRIORITY_ROW, i + PLAYER_INDEX_OFFSET);
    var role = ss.getRange(ROLE_ROW, i + PLAYER_INDEX_OFFSET).getValue();
    var name = ss.getRange(NAME_ROW, i + PLAYER_INDEX_OFFSET).getValue();
    switch(role)
    {
      case "Grandma":
      priocell.setValue(0);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue(2);
      ss.getRange(4, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(0, name);
      break;
      case "Escort":
      priocell.setValue(1);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(5, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(1, name);
      break;
      case "Godfather":
      priocell.setValue(2);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(7, ROLE_ORDER_OFFSET, 1, 2).setBackground('#ffcfc9');
      ScriptProperties.setProperty(2, name);
      break;
      case "Consigliere":
      priocell.setValue(3);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(8, ROLE_ORDER_OFFSET, 3, 2).setBackground('#ffcfc9');
      ScriptProperties.setProperty(3, name);
      break;
      case "Consort":
      priocell.setValue(3);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(8, ROLE_ORDER_OFFSET, 3, 2).setBackground('#ffcfc9');
      ScriptProperties.setProperty(3, name);
      break;
      case "Framer":
      priocell.setValue(3);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(8, ROLE_ORDER_OFFSET, 3, 2).setBackground('#ffcfc9');
      ScriptProperties.setProperty(3, name);
      break;
      case "Serial Killer":
      priocell.setValue(4);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(13, ROLE_ORDER_OFFSET, 1, 2).setBackground('#e6cff2');
      ScriptProperties.setProperty(4, name);
      break;
      case "Survivor":
      priocell.setValue(5);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue(1);
      ss.getRange(14, ROLE_ORDER_OFFSET, 1, 2).setBackground('#e6cff2');
      ScriptProperties.setProperty(5, name);
      break;
      case "Vigilante":
      priocell.setValue(6);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue(1);
      ss.getRange(16, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(6, name);
      break;
      case "Doctor":
      priocell.setValue(7);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(17, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(7, name);
      break;
      case "Nurse":
      priocell.setValue(8);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(18, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(8, name);
      break;
      case "Bodyguard":
      priocell.setValue(9);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(19, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(9, name);
      break;
      case "Investigator":
      priocell.setValue(10);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(20, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(10, name);
      break;
      case "Sheriff":
      priocell.setValue(11);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(21, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(11, name);
      break;
      case "Deputy":
      priocell.setValue(12);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(22, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(12, name);
      break;
      case "Lookout":
      priocell.setValue(13);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("Unlimited");
      ss.getRange(23, ROLE_ORDER_OFFSET, 1, 2).setBackground('#d4edbc');
      ScriptProperties.setProperty(13, name);
      break;
      default:
      priocell.setValue(14);
      ss.getRange(USES_ROW, i + PLAYER_INDEX_OFFSET).setValue("N/A");
    }
  }
}

function clearLastGame()
{
  // Clear past script properties
  ScriptProperties.deleteAllProperties();

  // Reset players
  for (var i = NUM_PLAYERS_CELL.getValue(); i < MAX_PLAYERS; i++)
  {
    ss.getRange(PRIORITY_ROW, i, MAX_PLAYERS * 2).setValule("");
    revivePlayer(i + PLAYER_INDEX_OFFSET);
  }

  // Recolor role order
  ss.getRange(4, ROLE_ORDER_OFFSET, 20, 2).setBackground('#888888');
  ss.getRange(11, ROLE_ORDER_OFFSET).setBackground('#000080');
}

function startNighttime()
{
  // Reset visit targets
  ss.getRange(VISIT_ROW, PLAYER_INDEX_OFFSET, 1, NUM_PLAYERS_CELL.getValue()).setValue("");

  // Highlight appropriate role in order list
  ss.getRange(4, ROLE_ORDER_OFFSET).activate();
  ScriptProperties.setProperty("turn", 0);
  
  // Update current player cell
  if (ScriptProperties.getProperty(0) != null)
  {
    CURRENT_PLAYER_CELL.setValue(ScriptProperties.getProperty(0));
  }
  else
  {
    CURRENT_PLAYER_CELL.setValue("Narrator");
  }

  // Debug turn index
  ss.getRange(6, 16).setValue(0);
}

function advanceNight()
{
  // Advance turn
  var turn = ScriptProperties.getProperty("turn");
  turn++;
  ScriptProperties.setProperty("turn", turn);

  // Update current player cell
  if (ScriptProperties.getProperty(turn) != null)
  {
    CURRENT_PLAYER_CELL.setValue(ScriptProperties.getProperty(turn));
    ScriptProperties.setProperty(CURRENT_PLAYER_CELL.getValue + "Ability", false);
  }
  else
  {
    CURRENT_PLAYER_CELL.setValue("Narrator");
  }

  // Highlight current role
  var highlightRow = turn + 4;
  if (turn > 1) highlightRow++;
  if (turn > 3) highlightRow += 4;
  if (turn > 5) highlightRow++;
  ss.getRange(highlightRow, ROLE_ORDER_OFFSET).activate();

  // Debug turn index
  ss.getRange(6, 16).setValue(turn);

  // Reminder to roleblock
  if (turn == 4)
  {
    var ui = SpreadsheetApp.getUI();
    ui.alert('Roleblock Reminder');
  }

  if (turn >= NUM_ROLES)
  {
    CURRENT_PLAYER_CELL.setValue("Nighttime over");
    endNighttime();
  }
}

function endNighttime()
{
  if (!checkValidVisits())
  {
    NIGHT_ERROR_CELL.setValue("Invalid visits");
    return;
  }
  NIGHT_ERROR_CELL.setValue("");

  // TODO: Calculate deaths, lookout reading
}

function activateAbility()
{
  var t = "" // TODO: acquire target
  ScriptProperties.setProperty(CURRENT_PLAYER_CELL.getValue + "Ability", true);
  ScriptProperties.setProperty(CURRENT_PLAYER_CELL.getValue() + "Target", t)
  // TODO: Give investigative info
}

function checkValidVisits()
{
  for(var i = 0; i < NUM_PLAYERS_CELL.getValue(); i++)
  {
    var val = ss.getRange(VISIT_ROW, i + PLAYER_INDEX_OFFSET).getValue();
    var role = ss.getRange(ROLE_ROW, i + PLAYER_INDEX_OFFSET).getValue();
    switch(role)
    {
      // Non-targeting roles
      case "Survivor":
      case "Executioner":
      case "Jester":
      case "Grandma":
        if (val != "")
        {
          VISIT_ERROR_CELL.setValue(role + " cannot target another player");
          VISIT_ERROR_CELL.setBackground('#800000');
          VISIT_ERROR_CELL.activate();
          return false;
        }
      break;
      // Targeting roles
      default:
        if (val == "")
        {
          VISIT_ERROR_CELL.setValue(role + " must target another player");
          VISIT_ERROR_CELL.setBackground('#800000');
          VISIT_ERROR_CELL.activate();
          return false;
        }
    }
  }
  // No invalid visits
  VISIT_ERROR_CELL.setValue("");
  VISIT_ERROR_CELL.setBackground('#666666');
  return true;
}

function onOpen()
{
  SpreadsheetApp.getUI()
    .createMenu('Custom Menu')
    .addItem('Show alert', 'showAlert')
    .addToUi();
}

function onEdit(e)
{
  var RANGE = e.range;
  //var RANGE = ss.getRange(8, 3)

  // Names Edited
  if (RANGE.getRow() == NAME_ROW)
  {
    var playerCount = 0;
    for (var i = 0; i < MAX_PLAYERS; i++)
    {
      var p = ss.getRange(NAME_ROW, i + PLAYER_INDEX_OFFSET);
      if (!p.isBlank())
      {
        playerCount++;
      }
    }
    NUM_PLAYERS_CELL.setValue(playerCount);
  }

  // Role Dropdowns Edited
  if (RANGE.getRow() == ROLE_ROW)
  {
    checkValidRoles();
  }

  // Alive Edited
  if (RANGE.getRow() == ALIVE_ROW)
  {
    var val = ss.getRange(RANGE.getRow(), RANGE.getColumn()).getValue();
    val ? revivePlayer(RANGE.getColumn()) : killPlayer(RANGE.getColumn());
  }
}

function killPlayer(c)
{
  // Ensure player is marked dead
  ss.getRange(ALIVE_ROW, c).setValue(false);

  // Strikethrough player name
  var textStyle = SpreadsheetApp.newTextStyle()
    .setStrikethrough(true)
    .build();
  ss.getRange(NAME_ROW, c).setTextStyle(textStyle);

  // Recolor player column
  var infoRange = ss.getRange(NAME_ROW, c, (PRIORITY_ROW - NAME_ROW) + 1);
  infoRange.setBackground('#800000');
  var nightRange = ss.getRange(PRIORITY_ROW + 1, c, MAX_PLAYERS * 2);
  nightRange.setBackground('#333333');
}

function revivePlayer(c)
{
  // Ensure player is marked alive
  ss.getRange(ALIVE_ROW, c).setValue(true);

  // Un-Strikethrough player name
  var textStyle = SpreadsheetApp.newTextStyle()
    .setStrikethrough(false)
    .build();
  ss.getRange(NAME_ROW, c).setTextStyle(textStyle);

  // Recolor player column
  var range = ss.getRange(NAME_ROW, c, (PRIORITY_ROW - NAME_ROW) + 1 + (MAX_PLAYERS) * 2);
  range.setBackground('#666666');
}

function checkValidRoles()
{
  var numPlayers = NUM_PLAYERS_CELL.getValue();
  var uRoles = new Set();
  for (let i = 0; i < numPlayers; i++)
  {
    uRoles.add(ss.getRange(ROLE_ROW, i + PLAYER_INDEX_OFFSET).getValue());
  }
  if (uRoles.size != numPlayers)
  {
    ROLE_ERROR_CELL.setValue("INVALID ROLES");
  }
  else
  {
    ROLE_ERROR_CELL.setValue("");
  }
}
