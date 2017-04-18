#!/usr/bin/env node

// Set up the command-line options
var program = require('commander');
program
  .arguments('<file>')
  .usage('<file> [options...]')
  .action(function(file) {
  })
  .parse(process.argv);

// Parse the incoming excel file
var filename = process.argv[2];
console.log('Processing ' + filename);
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile(filename);

// Static default values for the RCTC
var rctc_name = "Reportable Condition Trigger Codes (RCTC)",
    rctc_purpose = "Triggers for initiating decision support for electronic case report",
    rctc_oid = "2.16.840.1.114222.4.11.7508"

// Get RCTC header data from the 'Value Sets' tab
var worksheet = workbook.Sheets['Value Sets'];
var rctc_definition_version = worksheet['B5']['w'],
    rctc_effective_start_date = worksheet['B6']['w'];

// Get the code systems from the 'Value Sets' tab
var row_index = find_row_index('Code Systems',worksheet);
var rctc_code_systems = read_table(row_index + 1, worksheet);

// Get the value set details (that are not on the value set worksheets for some reason)
var row_index = find_row_index('Value Set List',worksheet);
var rctc_value_sets = read_table(row_index + 1, worksheet);

// Get the value sets from the other worksheets
var value_set_list = []
num_of_value_sets = workbook.SheetNames.length - 2;
for (var i=0; i < num_of_value_sets; i++) {
  var worksheet = workbook.Sheets[workbook.SheetNames[2 + i]];
  // Grab all the info from the tab
  var name = worksheet['B2']['w'],
      oid = worksheet['B3']['w'],
      type = worksheet['B4']['w'],
      definition_version = worksheet['B5']['w'],
      steward = "CSTE Steward",
      author = "CSTE Author",
      purpose_clinical_focus = worksheet['B8']['w'],
      purpose_data_element_scope = worksheet['B9']['w'],
      purpose_inclusion_criteria = worksheet['B10']['w'],
      purpose_exclusion_criteria = worksheet['B11']['w'],
      note = worksheet['B12']['w'],
      code_systems = worksheet['B13']['w'];
  // Grab the relevant information from the Value Sets tab for this value set
  // FIXME: This assumes the tabs are in the same order as the table
  var updated = rctc_value_sets[i]['updated_date'],
      status = rctc_value_sets[i]['status'];
  var grouping_list_index = find_row_index('Grouping List', worksheet);
  var grouping_list = read_table(grouping_list_index + 1, worksheet);
  var code_list_index = find_row_index('Code List', worksheet);
  var code_list = read_table(code_list_index + 1, worksheet);
  // Compose the hash for the value set
  var value_set = {
    name,
    oid,
    type,
    definition_version,
    steward,
    author,
    purpose_clinical_focus,
    purpose_data_element_scope,
    purpose_inclusion_criteria,
    purpose_exclusion_criteria,
    note,
    code_systems,
    updated,
    status,
    grouping_list,
    code_list
  };
  value_set_list.push(value_set);
}

// Compose the hash, given the RCTC header information and the value set list
var rctc_hash = {
  rctc_name,
  rctc_purpose,
  rctc_oid,
  rctc_definition_version,
  rctc_effective_start_date,
  rctc_code_systems,
  value_set_list
};

// Convert the hash to JSON
var rctc_json = JSON.stringify(rctc_hash, null, ' ');

// Convert the JSON to XML
// FIXME: Find a better program to create xml versions of the RCTC
var js2xmlparser = require("js2xmlparser");
var rctc_xml = js2xmlparser.parse("RCTC", rctc_hash);

// Output JSON, XML to file
// FIXME: Assumes that xls or xlsx only show up in the file type
var json_filename = filename.replace(/xlsx/g,"json").replace(/xls/g,"json")
var xml_filename = filename.replace(/xlsx/g,"xml").replace(/xls/g,"xml")
var fs = require('fs');
fs.writeFile(json_filename, rctc_json, function(err) {
  if(err) { return console.log(err); }
  console.log("The json file was saved: " + json_filename);
});
fs.writeFile(xml_filename, rctc_xml, function(err) {
  if(err) { return console.log(err); }
  console.log("The xml file was saved: " + xml_filename);
});

/*
Helper Functions below...
*/

// Returns the row of a given value in Column A for a worksheet
function find_row_index(value,worksheet){
  var range = XLSX.utils.decode_range(worksheet['!ref']); // get the range
  for(var R = range.s.r; R <= range.e.r; ++R) {
    var cellref = XLSX.utils.encode_cell({c:0, r:R}); // construct A1 reference for cell
    if(!worksheet[cellref]) continue; // if cell doesn't exist, move on
    var cell = worksheet[cellref];
    if(!(cell.t == 's' || cell.t == 'str')) continue; // skip if cell is not text
    if(cell.v === value) return R; // return the cell row index
  }
};

// Returns an array of values in a given row for a worksheet
function read_row_values(row, worksheet) {
  var row_array = [];
  var row_ref = parseInt(row);
  var range = XLSX.utils.decode_range(worksheet['!ref']); // get the range
  for(var C = range.s.c; C <= range.e.c; ++C) {
    var cellref = XLSX.utils.encode_cell({c:C, r:row_ref});
    if(!worksheet[cellref]) continue; // if cell doesn't exist, move on
    var cell = worksheet[cellref];
    row_array.push(cell.w);
  }
  return row_array;
};

// Given two arrays (one of keys, one of values), return an associative array
function arrays_to_hash (keys, values) {
  var hash = new Object();
  for(var i=0; i < values.length; i++){
    hash[keys[i]]=values[i];
  }
  return hash;
}

// Returns an array of hashes for a table within a worksheet, given the row of the headers and the worksheet
function read_table(row, worksheet) {
  var table = []
  var header_array = read_row_values(row, worksheet);
  // Clean the header array
  for (var i=0; i < header_array.length; i++){
    header_array[i] = header_array[i].replace(/ /g,"_").toLowerCase();
  };
  var row_ref = parseInt(row + 1);
  var range = XLSX.utils.decode_range(worksheet['!ref']); // get the range
  for(var R = row_ref; R <= range.e.r; ++R) {
    var cellref = XLSX.utils.encode_cell({c:0, r:R}); // construct A1 reference for cell
    if(!worksheet[cellref]) break; // if cell doesn't exist, stop
    var cell = worksheet[cellref];
    if(!(cell.t == 's' || cell.t == 'str')) break; // stop if cell is not text
    var values = read_row_values(R, worksheet);
    var hash = arrays_to_hash(header_array, values);
    table.push(hash);
  }
  return table;
};
