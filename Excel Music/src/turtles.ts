import {numberToLetter, lettersToNumber, getCellCoords, dynamicToVolume, expandRange} from './conversions';
import {isNote, isTurtle, isCell, isMultiNote, isDirChange, isDynamic} from './regex';
import {parseBrackets, processParsedBrackets} from './bracketsParse';
import * as Tone from 'tone';
import {piano} from './index';

const defaultSpeed = 160; // For when relative speeds are given
const defualtDynamic = 'mf' // For when a starting dynamic isn't given

/**
 * Checks selected sheet for cells that can be highlighted
 * @param sheet Used range of  the worksheet
 */
export function highlightSheet(sheet: Excel.Range): void {
    var sheetVals: any[][] = sheet.values;
    var rows: number = sheetVals.length;
    var cols: number = sheetVals[0].length;

    for (var row:number = 0; row < rows; row++){
        for (var col:number = 0; col < cols; col++){
            var value:string = sheetVals[row][col];
            // Highlight notes red
            if (isNote(value) || isMultiNote(value) && value.split(",").some(isNote)) {
                sheet.getCell(row,col).format.fill.color = "#FFada5";
            }
            // Highlight sustains a lighter red
            else if (value == "s" || value == "-" || value =="."){
                sheet.getCell(row,col).format.fill.color = "#FFd6d6";
            }
            // Highlight turtles green
            else if (isTurtle(value)) {
                sheet.getCell(row, col).format.fill.color = "#a8ffd0";
            }
            // Else remove highlighting
            else {
                sheet.getCell(row, col).format.fill.clear();
            }
        }
    }
}

/**
 * @param sheetVals the values in the used spreadsheet range
 * @return the beats per minute
 */
export function getBPM(): number {
    return defaultSpeed;
}

/**
 * Assuming a string isNote, returns the number note, octave and volume
 * @param cellValue the string written in the cell that isNote
 * @param currentOct the current octave at that point in the createNoteTimes process
 * @return [note, octave, volume]
 */
export function getNoteOctVol(cellValue: string, currentOct: number): [string, number, number] {
    var note: string;
    var octave: number;
    var volume: number = null;

    // Get note (and volume) from cell
    var cellNoteVol = cellValue.split(' ');

    // Process the volume if it is defined
    if (cellNoteVol.length == 2) {
        var cellVol = cellNoteVol[1];
        // process dynamic
        if (isDynamic(cellVol)) {
            volume = dynamicToVolume(cellVol);
        }
        else {
            volume = +cellVol;
        }
    }

    // get note (with octave)
    // octave defined in cell
    var cellNote = cellNoteVol[0];
    if (!isNaN(+cellNote[cellNote.length -1])) {
        octave = +cellNote[cellNote.length -1];
        note = cellNote;
    }
    // octave not defined in cell
    else {
        note = cellNote + currentOct;
        octave = currentOct;
    }
    return [note, octave, volume];
}

/**
 * Assuming a string isMultiNote, returns the number of notes
 * @param s a string
 * @return number of notes in multiNote
 */
export function countMultiNote(s: string): number {
    return s.split(',') //get each note
            .map(x=> x.trim().split(" ")[0]) // remove paddin then get note part
            .filter(function(x){return isNote(x)}) // remove if not a note
            .length;
}

/**
 * Takes a list of cell values and creates a list of time and note pairs for Tone Part playback
 * @param values List of notes cell contents
 * @param speedFactor Multipication factor for playback speed
 * @return list of time and note pairs for Tone Part playback
 */
export function createNoteTimes(values: [string, number][]): [[string, [string, string, number]][],number] {

    // find how many notes are defined
    var notesCount = 0;
    for(let valVol of values){
        value = valVol[0];
        if(isNote(value)){
            notesCount++;
        }
        if(isMultiNote(value)){
            notesCount = notesCount + countMultiNote(value);
        }
    }
    // Initialise an array of that many notes
    var noteSequence: [string, [string, string, number]][] = new Array(notesCount);

    var beatCount = 0; // how many cells through
    var noteCount = 0; // how many notes through
    var noteLength = 0; // for keeping track of note sustain
    // Current states
    var inRest = true // if the current value in the trace is a rest (else we're in a note)
    var currentStart: string; // start time of note currently in
    var currentNote: string; // note currently being played
    var currentVolume: number = dynamicToVolume('mf');
    // Per cell variables
    var volume: number;
    var value: string;
    var octave: number = 4;

    for (let valVol of values) {
        // volume = valVol[1]; 
        value = valVol[0];

        if(isNote(value)){

            // infer volume and octave from current values and cell values
            var noteOctVol = getNoteOctVol(value, octave);
            value = noteOctVol[0];
            octave = noteOctVol[1];
            if (noteOctVol[2] != null) {
                // Update the volume with value from cell
                volume = noteOctVol[2];
            }
            else {
                volume = currentVolume;
            }

            if(inRest){
                // Rest -> Note
                inRest = false;
            }
            else{
                // Note -> Note
                // end current note
                noteSequence[noteCount++] = [currentStart, [currentNote, "0:" + noteLength + ":0", currentVolume]];
            }
            currentStart = "0:" + beatCount + ":0"; // start new note
            currentNote = value;
            noteLength = 1;
            currentVolume = volume;
        }
        else if(value == null){ //rest
            if(!inRest){
                // Note -> Rest
                // end current note
                noteSequence[noteCount++] = [currentStart, [currentNote, "0:" + noteLength + ":0", currentVolume]];
                inRest = true;
            }
        }
        else if(value == 's' || value == '-'){
            // x -> x
            noteLength++;
        }
        else if(isMultiNote(value)){
            // Perform same actions as about but accounting for shorter time associated with each note
            var noteList = value.split(',')
                                .map(x => x.trim());
            var subdivisionLength = 1/(value.replace(/ /g,'')
                                            .split(',')
                                            .length);

            var subdivisionCount = 0;
            
            //now process all the scenarios of things than could be in the multinote
            for (let multiVal of noteList) {

                if(isNote(multiVal)){

                    var multiNoteOctVol = getNoteOctVol(multiVal, octave);
                    multiVal = multiNoteOctVol[0];
                    octave = multiNoteOctVol[1];
                    if (multiNoteOctVol[2] != null) {
                        volume = multiNoteOctVol[2];
                    }
                    else {
                        volume = currentVolume;
                    }

                    if(inRest){
                        // Rest -> Note
                        inRest = false;
                    }
                    else{
                        // Note -> Note
                        // end current note
                        noteSequence[noteCount++] = [currentStart, [currentNote, "0:" + noteLength + ":0", currentVolume]];
                    }
                    currentStart = "0:" + (beatCount+subdivisionCount*subdivisionLength) + ":0"; // start new note
                    currentNote = multiVal;
                    noteLength = subdivisionLength;
                }
                else if(multiVal == null){ //rest
                    if(!inRest){
                        // Note -> Rest
                        // end current note
                        noteSequence[noteCount++] = [currentStart, [currentNote, "0:" + noteLength + ":0", currentVolume]];
                        inRest = true;
                    }
                }
                else if(multiVal == 's' || multiVal == '-'){
                    // x -> x
                    noteLength += subdivisionLength;
                }
                currentVolume = volume;
                subdivisionCount++;
            }
        }
        beatCount++
    }
    // add note if we finished in a note
    if(!inRest){
        noteSequence[noteCount++] = [currentStart, [currentNote, "0:" + noteLength + ":0", currentVolume]];
    }
    return [noteSequence, beatCount];
}

/**
 * Takes a list of Cell Values and plays a Tone sequence
 * @param values List of notes as strings e.g. ['A4','A5']
 * @param speedFactor Multipication factor for playback speed
 * @param repeats number of repeats in turtle -> how many times sequence is played
 */
export function playSequence(values: [string, number][], speedFactor: number =1, repeats: number =0): void {
    // Convert cells in turtles path to note timings
    var [noteTimes, beatsLength]: [[string, [string, string, number]][],number] = createNoteTimes(values);
    
    // If Piano sample is not being used can use built in synth
    var polySynth = new Tone.PolySynth(4, Tone.Synth, {
        "volume" : -4,
        "oscillator" : {
            "partials" : [1, 2, 1],
        },
        "portamento" : 0.05
    }).toMaster();

    console.log(noteTimes);

    // Create Tone.Part
    var turtlePart = new Tone.Part(function(time: string, note: [string, string, number]){
        piano.triggerAttackRelease(note[0], note[1], time, note[2]);
    }, noteTimes).start();

    turtlePart.loop = true;
    turtlePart.loopEnd = "0:" + beatsLength + ":0"; // time of a single loop unit
    // If not looping indefiniely, declare stopping time
    if (repeats>0){
        turtlePart = turtlePart.stop("0:" + (repeats*beatsLength/speedFactor) + ":0");
    }

    turtlePart.humanize = false;
    turtlePart.playbackRate = speedFactor;
}

/**
 * Next direction of a turtle given current direction and instruction
 * @param current current compass direction being faced
 * @param move next way to turn/look
 * @return direction facing after following instruction
 */
export function dirChange(current: string, move: string): string {
    current = current.toLowerCase();
    // If direction given is absolute, return that
    if (RegExp(/^(n|e|s|w)$/).test(move)) {
        return move;
    }
    // If relative, update as required
    else {
        if (move == 'r') {
            switch (current) {
                case 'n': return 'e';
                case 'e': return 's';
                case 's': return 'w';
                case 'w': return 'n';
            }
        }
        else {
            switch (current) {
                case 'n': return 'w';
                case 'e': return 'n';
                case 's': return 'e';
                case 'w': return 's';
            }
        }
    }    
}

/**
 * Given current coordinates and direction, return coordinates after step forwards
 * @param current current coordinates
 * @param dir compass direction turtle is facing
 * @return new coordinates of turtle
 */
export function move(current: [number, number], dir: string): [number, number] {
    // Can't go west of column A or north of row 1
    switch (dir) {
        case 'n': return [Math.max(current[0]-1,0), current[1]];
        case 'e': return [current[0], current[1]+1];
        case 's': return [current[0]+1, current[1]];
        case 'w': return [current[0], Math.max(current[1]-1,0)];
    } 
}

/**
 * Given current coordinates and direction, return coordinates after step forwards
 * @param start
 * @param moves
 * @param sheetVals
 * @return [notes, trace] where notes is a list of cell contents and trace is a list of coordinates
 */
export function getTurtleSequence(start: string, moves: string[], sheetVals: any[][]): [[string, number][],[number,number][]] {
    console.log(start);

    var startCoords: [number, number] = getCellCoords(start);
    var volume: number = dynamicToVolume(defualtDynamic);

    var sheetValsRows = sheetVals.length;
    var sheetValsCols = sheetVals[0].length;

    // notes are stored in the format [note, volume] at this stage
    var notes: [string, number][] = [[sheetVals[startCoords[1]][startCoords[0]], volume]];

    // place in starting cell facing north
    var dir: string = 'n';
    var pos: [number, number] = [startCoords[1],startCoords[0]];

    // trace for tracking where the turtle is and highlighting - post diss addition
    var trace = [pos];

    for (let entry of moves) {
        if (isDirChange(entry)) {
            if (entry.length > 1) {
                var j: number = +entry.substring(1);
                var rotation = entry.substring(0,1);
                // Turn the number of times put after the direction change
                while (j > 0) {
                    dir = dirChange(dir, rotation);
                    j = j - 1;
                }
            }
            else {
                // If number of times not specified, do it once
                dir = dirChange(dir, entry);
            }
        }
        else if (entry.substring(0,1) == 'j' || entry.substring(0,1) == 'J') {
            // Jump
            var jumpInstructions = entry.substring(1);
            if (isCell(jumpInstructions)) {
                // Move to given cell
                var newCoords = getCellCoords(jumpInstructions);
                pos = [newCoords[1], newCoords[0]];
            }
            else {
                // relative Jump
                var regex = /(\+|-)[0-9]+/g; // regex for one part of the relative jump instruction
                var movements = jumpInstructions.match(regex) // find two direction instrucitons of jump
                                                .map(x => +x);
                pos = [pos[0] + movements[1], pos[1] + movements[0]];
            }
            notes.push([sheetVals[pos[0]][pos[1]],volume]);
            trace.push(pos);
        }
        else if (isDynamic(entry)) {
            // No longer supported
            volume = dynamicToVolume(entry);
            console.log("Dynamic in instructions - outdated. Dynamics should go in cell")
        }
        else {
            // move
            var steps : number;
            if (entry == "m") {
                steps = 1;
            }
            else if (entry.substring(1) == "*") {
                // move as far as there is data
                // get array it's moving into
                console.log("*");
                console.log(pos, dir);
                var arr: any[];
                // Get the array in the direction it is facing
                // part of array depends which was the turtle is facing.
                if (dir == 's') {
                    arr = sheetVals.map(function(value,index) {
                                        return value[pos[1]]; 
                                        }) // get column cell from each row
                                    .slice(pos[0]+1);
                }
                else if (dir == 'n') {
                    arr = sheetVals.map(function(value,index) {
                                        return value[pos[1]]; 
                                    })
                                    .slice(0,pos[0])
                                    .reverse();
                }
                else if (dir == 'e') {
                    arr = sheetVals[pos[0]].slice(pos[1]+1);
                }
                else {
                    // w
                    arr = sheetVals[pos[0]].slice(0,pos[1]);
                }
                
                // find last element that is a note
                steps = arr.length - arr.slice() //to get shallow copy
                                        .reverse()
                                        .findIndex(x => isNote(x) || x == "s" || x == "-" || x=="." || isMultiNote(x));
                // Potential change: Make sure that if there are explicit rests, they occur at the end. - actually if it end on a note or sustain this doesn't quite wrap things up
            }
            else {
                // integer part of the move defined
                steps = +entry.substring(1); 
            }
            console.log("steps: " + steps);
            var i: number;
            var sheetVal;
            // take the defined (or infered) number of steps
            for (i = 0; i < steps; i++) {
                pos = move(pos, dir);
                if (pos[0] >= sheetValsRows || pos[1] >= sheetValsCols) {
                    // if moving out of defined spreadsheet range
                    sheetVal = null;
                }
                else{
                    sheetVal = sheetVals[pos[0]][pos[1]];
                    if (sheetVal == "") {
                        sheetVal = null;
                    }
                }
                notes.push([sheetVal, volume]);
                trace.push(pos);
            }
        }

    }
    return [notes,trace];
}

/**
 * Runs a turtle that plays the notes in the cells it passes through
 * @param instructions Instructions as defined by the user in the cell: !turtle(<instrutions>)
 * @param sheetVals The values in the used spreadsheet range
 */
export function turtle(instructions: string, sheetVals: any[][], sheet: Excel.Range): void {

    var instructionsArray: string[] = instructions.split(',');

    if (isCell(instructionsArray[0])) {
        var notes: [string, number][];
        var trace: [number, number][];
        // Instead of instructions, end cell given. No longer advertised as a feature
        if (isCell(instructionsArray[1])){
            var rangeStart = getCellCoords(instructionsArray[0]);
            var rangeEnd = getCellCoords(instructionsArray[1]);
            notes = [].concat.apply([], sheetVals.slice(rangeStart[1], rangeEnd[1]+1)
                .map(function(arr) { 
                    return arr.slice(rangeStart[0], rangeEnd[0]+1).map(function(x) {
                            return [x,dynamicToVolume('mf')];
                    });
                })
            );
        }
        else{
            // Start cell and movement instructions
            var start: string = instructionsArray[0];
            var moves: string[] = processParsedBrackets(parseBrackets(instructionsArray[1])).split(" ");
            [notes, trace] = getTurtleSequence(start, moves, sheetVals);
        }
        var speedFactor: number = 1;
        var repeats: number = 0;
        if (instructionsArray.length > 2){
            // Speed given
            speedFactor = eval(instructionsArray[2]);
            if (speedFactor > 10) {
                // <10 currently supports relative speeds still
                speedFactor = speedFactor / defaultSpeed;
            }
            if (instructionsArray.length > 3){
                // Repeats given
                repeats = +instructionsArray[3].replace(/\s/g, "");
            }
        }
        // console.log(notes);
        playSequence(notes, speedFactor, repeats);
        console.log(trace);
        // Tracing
        // console.log("tracing");
        // for (let cellCoords of trace){
        //     console.log(cellCoords);
        //     var cellBoarders = sheet.getCell(cellCoords[0],cellCoords[1]).format.borders;
            // cellBoarders.getItem('EdgeBottom').weight = 'Thick';
            // cellBoarders.getItem('EdgeBottom').color = 'Green';
            // cellBoarders.getItem('EdgeRight').weight = 'Thick';
            // cellBoarders.getItem('EdgeRight').color = 'Green';
            // cellBoarders.getItem('EdgeLeft').weight = 'Thick';
            // cellBoarders.getItem('EdgeLeft').color = 'Green';
            // cellBoarders.getItem('EdgeTop').weight = 'Thick';
            // cellBoarders.getItem('EdgeTop').color = 'Green';
        // }
        // console.log('Done Tracing');
    }
    else {
        // mutliple turtles
        var turtlesStarts = expandRange(instructionsArray[0].replace(/\s/g, "")); // list of starting notes
        var moves: string[] = processParsedBrackets(parseBrackets(instructionsArray[1])).trim().split(" ");
        if (instructionsArray.length > 2){
            // Defined Speed
            speedFactor = +instructionsArray[2].replace(/\s/g, "");
            if (speedFactor > 10) {
                // Absolute Speed
                speedFactor = speedFactor / defaultSpeed;
            }
            if (instructionsArray.length > 3){
                // Repeats given
                repeats = +instructionsArray[3].replace(/\s/g, "");
            }
        }
        // For each turtle, play that turtle
        for (let turtleStart of turtlesStarts){
            [notes, trace] = getTurtleSequence(turtleStart, moves, sheetVals);
            playSequence(notes, speedFactor, repeats);
        }
    }
}

/**
 * Finds all turtle declarations in the spreadsheet and runs them
 * @param sheetVals values in the spreadsheet
 */
export function runTurtles(sheet: Excel.Range): void {
    var sheetVals:any[][] = sheet.values;
    var rows: number = sheetVals.length;
    var cols: number = sheetVals[0].length;
    // Make list of active turtles
    var live_turtles = document.createElement('ul');
    live_turtles.setAttribute('class','ms-List');

    var row: number, col: number;
    for (row = 0; row < rows; row++) {
        for (col = 0; col < cols; col++) {
            var value = sheetVals[row][col];
            if (isTurtle(value)) {
                var  instructions = value.substring(8, value.length - 1);
                turtle(instructions, sheetVals,sheet);
                // Add to list of active turtles
                var live_turtle = document.createElement('li');
                live_turtle.appendChild(document.createTextNode(numberToLetter(col).toUpperCase()+(row+1)));
                live_turtles.appendChild(live_turtle);
            }
        }
    }
    document.getElementById('live_turtles').appendChild(live_turtles);
}