import { Distance } from "tonal";

declare var CustomFunctionMappings;

var re_isNote = new RegExp('^[A-G](#|b|)?[1-9]?( (0(\.\[0-9]+)?|1(\.0)?|ppp|pp|p|mp|mf|f|ff|fff))?$');
function isNote(val: string): boolean {
    return re_isNote.test(val);
}

function isMultiNote(s: string): boolean {
    if (typeof s != 'string') {
        return false;
    }
    if (!(s.includes(','))) {
        return false;
    }
    var arr = s.split(',');
    for (let val of arr) {
        val = val.trim();
        if (!isNote(val) && !(val=="") && !(val=='s') && !(val=='-')){
            return false
        }
    }
    return true;
}

/**
 * If a string is musical dynamic / volume
 * @param s a string
 * @return if it is a dynamic
 */
export function isDynamic(s: string) : boolean {
    s = s.toLowerCase();
    return RegExp(/^(ppp|pp|p|mp|mf|f|ff|fff)$/).test(s);
}

function modulate(cell: string, interval: string): string {
  if (isNote(cell)) {
    var splitCell = cell.split(" ");
    var dynamic = (splitCell.length > 1) ? " " + splitCell[1] : "";
    return Distance.transpose(splitCell[0], interval).toString() + dynamic;
    // TODO: add dynamic back again
  }
  if (isMultiNote(cell)) {
    var notes = cell.split(",").map(x => x.trim());
    for (var i=0; i<notes.length; i++) {
      var note = notes[i];
      if (isNote(note)) {
        notes[i] = modulate(note, interval);
      }
    }
    return notes.toString();
  }
  else {
    return cell;
  }
}

function turtle(start_cell: string, instructions: string, speed: number = null, loops: number = null): string {
  console.log(start_cell);
  if (loops == null ) {
    if (speed == null) {
      return "!turtle(" + start_cell + ", " + instructions + ")";
    }
    return "!turtle(" + start_cell + ", " + instructions + ", " + speed + ")";
  }
  return "!turtle(" + start_cell + ", " + instructions + ", " + speed + ", " + loops + ")";
}

function increment(incrementBy: number, callback) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = () => {
    clearInterval(timer);
  };
}

CustomFunctionMappings.MODULATE = modulate;
CustomFunctionMappings.TURTLE = turtle;
CustomFunctionMappings.INCREMENT = increment;
