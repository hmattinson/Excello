/**
 * Gives the column heading of a number e.g. 1->A, 27->AA
 * @param x number assuming integer
 * @return x base 26 using the letters of the alphabet
 */
export function numberToLetter(x: number): string {
    var leastSignificantColumn = 'abcdefghijklmnopqrstuvwxyz'[x % 26];
    if (x >= 26) {
        return numberToLetter((x/26) - 1) +  leastSignificantColumn;
    }
    else {
        return leastSignificantColumn;
    }
    // return (x >= 26 ? numberToLetter((x / 26 >> 0) - 1) : '') +  'abcdefghijklmnopqrstuvwxyz'[x % 26 >> 0];
}

/**
 * Gives the index of the column of given letters
 * @param letters column e.g. AB
 * @return index of this column
 */
export function lettersToNumber(letters: string): number {
    var num: number = 0;
    const len: number = letters.length;
    letters = letters.toUpperCase();
    // sum(num(c[i]) * 26^i)
    for (var i = 0; i < len; i++) {
        num += (letters.charCodeAt(i) - 64) * Math.pow(26, len - i - 1);
    }
    return num;
}

/**
 * Gives the index coordinates of a cell using Excel coordinates (column, row)
 * @param letters cell position e.g. B1
 * @return coordinates with 0 indexing
 */
export function getCellCoords(battleship: string): [number, number] {
    // split into column and row
    var x = battleship.match(/[a-zA-Z]+|[0-9]+/g);
    return [lettersToNumber(x[0]) - 1, +x[1] - 1]; // minuses becuase of different indexings
}

/**
 * Given current a dynamic, return volume in range [0,1]
 * @param dynamic dynamic marking
 * @return number in [0,1]
 */
export function dynamicToVolume(dynamic: string): number {
    switch (dynamic) {
        case 'ppp': return 0.125;
        case 'pp': return 0.25;
        case 'p': return 0.375;
        case 'mp': return 0.5;
        case 'mf': return 0.625;
        case 'f': return 0.75;
        case 'ff': return 0.875;
        case 'fff': return 1;
    } 
}

/**
 * takes a start and end cell and gives addresses of cells between
 * @param range e.g. B1:B10
 * @return list of addresses in range (inclusive)
 */
export function expandRange(range: string): string[] {
    var [start, end] = range.split(":");
    var startCoords = getCellCoords(start);
    var endCoords = getCellCoords(end);
    var colChange = endCoords[0] - startCoords[0];
    var rowChange = endCoords[1] - startCoords[1];
    var startSplit = start.match(/[a-z]+|[^a-z]+/gi);
    var endSplit = end.match(/[a-z]+|[^a-z]+/gi);
    if (rowChange!=0) {
        var col: string = startSplit[0];
        var cells: string[] = []
        for (var i = Math.min(+startSplit[1], +endSplit[1]); i <= Math.max(+startSplit[1], +endSplit[1]); i++) {
            cells.push(col + i);
        }
    }
    else if (colChange != 0) {
        var row: string = startSplit[1];
        var cells : string[] = [];
        for (var i=lettersToNumber(startSplit[0])-1; i<=lettersToNumber(endSplit[0])-1; i++) {
            cells.push(numberToLetter(i) + row);
        }
    }
    return cells;
}