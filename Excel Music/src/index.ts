import * as OfficeHelpers from '@microsoft/office-js-helpers';
import * as Tone from 'tone';
import {Chord, Note} from 'tonal';
import {isTurtle} from './regex';
import {insertChord, addChordOctave} from './chords';
import {parseBrackets, processParsedBrackets} from './bracketsParse';
import {runTurtles, highlightSheet, getBPM} from './turtles';

$("#run").hide();
$("#run").show(); // may want to use this to let sounds load first

$("#refresh").click(() => tryCatch(refresh));
$("#run").click(() => tryCatch(run));
$("#stop").click(() => tryCatch(stop));
$("#toggle").click(() => tryCatch(toggle));
$("#insertChord").click(() => tryCatch(insertChord));

export var piano: Tone.Sampler;

/**
 * Run when play button pressed. Starts playback of music
 */
async function run() {
    await Excel.run(async (context) => {

        //  Colour button green
        document.getElementById("run").style.background='#A8FFD0';

        // Get values from the sheet selected in the drop down
        var sheetSelectHTMLElement = (document.getElementById("sheet_select")) as HTMLSelectElement;
        var sheetChoice = sheetSelectHTMLElement.options[sheetSelectHTMLElement.selectedIndex].value;
        const sheet: Excel.Range = context.workbook.worksheets.getItem(sheetChoice).getUsedRange();
        sheet.load('values');

        await context.sync();

        highlightSheet(sheet);

        // In order to remove the notes from the previous play from the transport, new audio context
        Tone.context.close();
        Tone.context = new AudioContext();
        Tone.Transport.bpm.value = getBPM();

        // Clear current live turtles
        var turtle_list_div = document.getElementById('live_turtles');
        while (turtle_list_div.firstChild) {
            turtle_list_div.removeChild(turtle_list_div.firstChild);
        }

        // Load piano samples and build sampler
        piano = new Tone.Sampler({
            'A0' : 'A0.[mp3|ogg]',
            'C1' : 'C1.[mp3|ogg]',
            'D#1' : 'Ds1.[mp3|ogg]',
            'F#1' : 'Fs1.[mp3|ogg]',
            'A1' : 'A1.[mp3|ogg]',
            'C2' : 'C2.[mp3|ogg]',
            'D#2' : 'Ds2.[mp3|ogg]',
            'F#2' : 'Fs2.[mp3|ogg]',
            'A2' : 'A2.[mp3|ogg]',
            'C3' : 'C3.[mp3|ogg]',
            'D#3' : 'Ds3.[mp3|ogg]',
            'F#3' : 'Fs3.[mp3|ogg]',
            'A3' : 'A3.[mp3|ogg]',
            'C4' : 'C4.[mp3|ogg]',
            'D#4' : 'Ds4.[mp3|ogg]',
            'F#4' : 'Fs4.[mp3|ogg]',
            'A4' : 'A4.[mp3|ogg]',
            'C5' : 'C5.[mp3|ogg]',
            'D#5' : 'Ds5.[mp3|ogg]',
            'F#5' : 'Fs5.[mp3|ogg]',
            'A5' : 'A5.[mp3|ogg]',
            'C6' : 'C6.[mp3|ogg]',
            'D#6' : 'Ds6.[mp3|ogg]',
            'F#6' : 'Fs6.[mp3|ogg]',
            'A6' : 'A6.[mp3|ogg]',
            'C7' : 'C7.[mp3|ogg]',
            'D#7' : 'Ds7.[mp3|ogg]',
            'F#7' : 'Fs7.[mp3|ogg]',
            'A7' : 'A7.[mp3|ogg]',
            'C8' : 'C8.[mp3|ogg]'
            }, {
            'release' : 1,
            'baseUrl' : '../assets/samples/salamander/',
            'onload': function() {
                runTurtles(sheet);
                Tone.Transport.start("+0.1");
            }
        }).toMaster();

        // // If using synth not piano samples
        // runTurtles(sheet.values);
        // Tone.Transport.start("+0.2");
    });
}

/**
 * Stops music playback
 */
async function stop() {
    await Excel.run(async (context) => {
        document.getElementById("run").style.removeProperty("background-color");
        Tone.context.close(); // close audio context
    });
}

/**
 * refreshes sheets in the dropdown based on sheets that are in workbook
 */
async function refresh() {
    await Excel.run(async (context) => {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        // Empty and rebuild the current options in dropdown
        $("#sheet_select").empty()
        for (var i in sheets.items) {
            var name = sheets.items[i].name
            // Add option to select
            $('#sheet_select').append($('<option>', {
                value: name,
                text: name
            }));
        }
    });
}

/**
 * If a turtle cell is selected, it will toggle the activation of the cell
 */
async function toggle() {
    await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("values");
        selectedRange.load("address");
        await context.sync();
        var newValues = selectedRange.values;
        // add/remove ! as required
        for (var row=0; row<newValues.length; row++) {
            for (var col=0; col<newValues[0].length; col++) {
                var val = newValues[row][col]
                // Current Turtle definition includes the !
                if (isTurtle(val)) {
                    // Remove !
                    newValues[row][col] = val.substring(1);
                }
                // Add ! if that makes it a turtle
                else if (isTurtle("!" + val)) {
                    newValues[row][col] = "!" + val;
                }
            }
        }
        // Propogate changes to grid
        const selectedSheet = context.workbook.worksheets.getActiveWorksheet();
        selectedSheet.getRange(selectedRange.address.split('!')[1]).values = newValues;
    });
}

/** Default helper for invoking an action and handling errors. */
export async function tryCatch(callback) {
    try {
        await callback();
    }
    catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    }
}