# Excello

This project took Microsoft's Excel, one of the most omnipresent pieces of software, and created a system for music composition and performance within it - Excello.

## Set up

### For Users

* Download the manifest.xml file from the Excel Music folder
* Download the Excel Music spreadsheet (.xlsx)
* Open Excel (a version that supports Add-ins - the online version works fine)
* (Upload and) open the spreadsheet.
* Insert > Office Add-ins > Upload My Add-in
* Find and upload the manifest file
* Excello will now be available to open from the home tab.

### For Developers

- Clone repo
- Sort out node stuff
- run from within the repo (npm run start:web)
- Open included spreadsheet in Excel Online
- Add the local_manifest from this project

This is helpful: https://docs.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial?tutorial-step=1

## Using Excello

The Excel Music spreadsheet includes a tutorial and demos.

#### Cells

Notes are defined in the cells of the spreadsheet. Using "scientific pitch notation" notes are defined by note name and octave number e.g. "A4". The first note of the octave is C. If the octave number is omitted, the octave of the previously played note is used. **Volume** is defined using either a dynamic marking or a number in the range [0,1] separated from the note by a space. e.g. "A4 0.8", "G pp".

| Dynamic Marking | Volume |
| --------------- | ------ |
| ppp             | 0.125  |
| pp              | 0.25   |
| p               | 0.375  |
| mp              | 0.5    |
| mf              | 0.625  |
| f               | 0.75   |
| ff              | 0.875  |
| fff             | 1      |

Notes are **sustained** by placing "-" in the next cell in the path.

A **rest** is simply an empty cell.

It is also possible to **subdivide** cells using commas e.g. "C3,D3,E3" would create triplets or "s,s,,F4" which would sustain the previous notes for an additional 2 quarter measures, a quarter measure rest and then F4 for the last quarter of the beat.

In most musical interfaces one axis is time (and the other is normally pitch). In this case we wanted to explore the use of both axes as just space - ideomatic to Excel. As a result the user is free to arrange the data as they wish.

#### Turtles

Notes are played by defining turtles to navigate the spreadsheet. Turtles are defined as follow

```
!turtle(<Starting Cell>, <Movement>, <Speed>, <Number of loops>)
```
The "!" dictates that the turtle will be activated when the play button is pressed.

##### Starting Cell

e.g. B2. As with normal Excel formulae, you provide a reference to a cell by using the battleship style coordinates of the cell. This is where the turtle will start. This cell will also be played and forms the first cell in the path of the turtle. 

To define multiple turtles following identical paths but starting from adjacent cells, rather than writing a cell per turtle you can define a range of starting cells:

```
!turtle(E6:E11, r m2 r m2 l m3, 1)
```

##### Instructions

The following expressions can be used:

* Relative turning:
  * r: turn right
  * l: turn left
* Absolute turning:
  * n: turn to face up/north
  * e: turn to face right/east
  * s: turn to face down/south
  * w: turn to face left/west
* movement:
  * m: move the number of cells specified forward e.g. m3
    * m* can be used to move forward until the last note/sustain in that path
* Jumps j:
  * absolute jumps e.g. 'jB2' - jump to a cell (that cell will also be played)
  * relative jumps e.g. 'j-14+4' (jump 14 cells left and 4 cells down) - two numbers each with an associated direction, first indicated how many cells right to move, the next how many cells down.

The turtle starts facing north by default. 

Just like "r2" can be written instead of "r r", the same idea can be used for larger parts of instructions. Using parenthesis, instruction can be repeated and nested.

```
!turtle(A1, (r m3)4)
```

This example is equivalent to the movement instructions "m3 r m3 r m3 r m3 r" and has the turtle follow a square

##### Speed

Default 160.

This is cells per minute that the turtle will move at. This used to be a relative amount so currently values less that 10 will be multiplied by 10 to maintain backwards compatibility. 

##### Number of loops

Default 0 - which creates an infinite loop

The number of times the turtle will travel through the path defined. If left blank or defined as 0, it will loop infinitely.

#### Adding Chords

Rather than working out each note for a chord and typing them in, in the window in the side you can select a chord and insert it into the spreadsheet. A chord has the following properties:

* Tonality of the chord e.g. C or F#
* Type of chord e.g. maj7, 9sus4
* Inversion of the chord - which number note in the chord is used to start
* Octave of the first note in the chord

### Excel Formulae

The Custom_Functions folder contains the implementation of two custom functions that can be loaded into Excel to help with the composition process. These are not hosted on an online version as they are not crucial for using Excel. If you wish to add them to your spreadsheet you will have to run this file locally as with the main Excello code. 

###### EXCELLO.TURTLE

This takes the 4 arguments as required for the turtle (last 2 optional) and creates the turtle definition. This is particularly useful for referring to a global speed variable. 

###### EXCELLO.TRANSPOSE

This allows the tonal library's transpose function to be used. The first argument of the function is any cell (it can include dynamics be multi-note). The second is an interval as defined in the tonal library (<https://github.com/danigb/tonal/tree/master/packages/interval#module_Interval>).

## MIDI converter

Included in this repo is a Python notebook for converting MIDI files into csv files that can be played in Excello. However, some of these can take some time to play on Excel, especially if they are particularly long or not running locally. 

