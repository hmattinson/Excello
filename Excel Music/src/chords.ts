import {Chord, Note} from 'tonal';
import {getCellCoords, numberToLetter} from '../src/conversions';

/**
 * Reads the selected chord info from the menu and inserts the chord notes into the cells
 */
export async function insertChord() {
    await Excel.run(async (context) => {

        const selectedRange = context.workbook.getSelectedRange();
        const selectedSheet = context.workbook.worksheets.getActiveWorksheet();
        selectedRange.load("address");
        selectedRange.load("rowCount");
        selectedRange.load("columnCount");

        // get info from the selects
        // Note
        var chordNoteHTMLElement = (document.getElementById("chord_note")) as HTMLSelectElement;
        var chordNote = chordNoteHTMLElement.options[chordNoteHTMLElement.selectedIndex].value;
        // Type
        var chordTypeHTMLElement = (document.getElementById("chord_type")) as HTMLSelectElement;
        var chordType = chordTypeHTMLElement.options[chordTypeHTMLElement.selectedIndex].value;
        // Octave
        var chordOctaveHTMLElement = (document.getElementById("octave")) as HTMLSelectElement;
        var chordOctave = chordOctaveHTMLElement.options[chordOctaveHTMLElement.selectedIndex].value;
        // Inversion
        var chordInversionHTMLElement = (document.getElementById("inversion")) as HTMLSelectElement;
        var chordInversion:number = +chordInversionHTMLElement.options[chordInversionHTMLElement.selectedIndex].value;

        // get notes of defined chord
        var chordNotes = Chord.notes(chordNote, chordType).map(x => Note.simplify(x));
        // invert the chord as required
        while(chordInversion--)chordNotes.push(chordNotes.shift());
        // Add the correct octave
        chordNotes = addChordOctave(chordNotes, chordOctave);
        
        await context.sync();

        // find cell for start of where chord will be inputted
        var selectedRangeStart = selectedRange.address.split('!')[1].split(':')[0];
        var selectedRangeStartCoords  = getCellCoords(selectedRangeStart);
        // if the chord should be inserted vertically
        var vertical: boolean = (selectedRange.rowCount - selectedRange.columnCount) >= 0;

        if (vertical) {
            // find the end cell for where the chord will be inputted
            var inputRangeEndCell = numberToLetter(selectedRangeStartCoords[0]) + (selectedRangeStartCoords[1]+chordNotes.length);
            var inputRange = selectedRangeStart + ':' + inputRangeEndCell;
            // Reverse so the order is higher pitch on top
            selectedSheet.getRange(inputRange).values = chordNotes.reverse().map(x => [x]);
        }else {
            var inputRangeEndCell = numberToLetter(selectedRangeStartCoords[0] + chordNotes.length-1) + (selectedRangeStartCoords[1]+1);
            var inputRange = selectedRangeStart + ':' + inputRangeEndCell;
            selectedSheet.getRange(inputRange).values = [chordNotes];
        }
        
    }).catch(errorHandlerFunction);
}

/**
 * Given a list of notes and a starting octave, applies octave numbers to the notes assuming always ascending
 * @param chordNotes List of notes without octave as strings e.g. ['C','E','G','Bb','D']
 * @param startingOctave the octave number of the first note in the list e.g. 4
 * @return The notes with octave numbers applied e.g. ['C4','E4','G4','Bb4','D5']
 */
export function addChordOctave(chordNotes: string[], startingOctave: string): string[]{
    var noteOrder = {
        'C': 1,
        'C#': 2,
        'Db': 2,
        'D': 3,
        'D#': 4,
        'Eb': 4,
        'E': 5,
        'F': 6,
        'F#': 7,
        'Gb': 7,
        'G': 8,
        'G#': 9,
        'Ab': 9,
        'A': 10,
        'A#': 11,
        'Bb': 11,
        'B': 12
      }; // used to check when an octave boundary has been passed
    var octave: number = +startingOctave; // octave won't have changed for first note
    var octavedNotes: string[] = new Array(chordNotes.length); // array that will be returned
    octavedNotes[0] = chordNotes[0] + startingOctave;
    var previousNote = chordNotes[0]
    for (var i=1; i<chordNotes.length; i++) {
        var note = chordNotes[i];
        // Check if we have looped around the octave
        if (noteOrder[note] <= noteOrder[previousNote]) {
            // new octave
            octave++;
        }
        octavedNotes[i] = note + octave;
        previousNote = note;
    }

    return octavedNotes;
}

function errorHandlerFunction(err) {
    console.log(err);
}