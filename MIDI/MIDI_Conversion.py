#!/usr/bin/env python
# coding: utf-8

# # MIDI to CSV Converter

# In[657]:


import mido
import csv
import os


# In[178]:


def cprint(s,b):
    if b:
        print(s)


# In[273]:


def get_active_notes(mid):
    active_notes = {}
    tracks = mid.tracks
    num_tracks = len(tracks)
    all_notes = num_tracks * [None]
    for i in range(0, num_tracks):
        track = tracks[i]
        time = 0
        all_notes[i] = []
        for msg in track:
            msg_dict = msg.dict()
            time += msg_dict['time']
            if msg.type == 'note_on' or msg.type == 'note_off':
                vel = msg_dict['velocity']
                if vel > 0 and msg.type == 'note_on':
                    # Using a list for the active notes becuase note 71 in io.mid was definied twice at once
                    if active_notes.has_key(msg_dict['note']):
                        active_notes[msg_dict['note']].append({'time':time,'velocity':vel})
                    else:
                        active_notes[msg_dict['note']] = [{'time':time, 'velocity': vel}]
                elif vel == 0 or msg.type == 'note_off':
                    note = msg_dict['note']
                    if len(active_notes[note])>0:
                        start_msg = active_notes[note].pop()
                        new_note = {'note': note, 'start': start_msg['time'],
                                    'end': time, 'velocity': start_msg['velocity']}
                        all_notes[i].append(new_note)
    return all_notes


# In[44]:


def create_streams(all_notes):
    streams = []
    for notes in all_notes:
        while notes != []:
            stream = []
            vel = 0
            current_end = 0
            for note in notes:
                if note['start'] >= current_end:
                    if note['velocity'] != vel:
                        vel = note['velocity']
                    else:
                        del note['velocity']
                    stream.append(note)
                    current_end = note['end']
            streams.append(stream)
            for note in stream:
                notes.remove(note)
    return streams


# In[645]:


midiNotes = ['C','C#','D','D#','E','F','F#','G','G#','A','A#','B']
def midi2str(midi):
    return midiNotes[midi%12] + str(midi/12 -1)


# In[653]:


def streams_to_cells(streams, speed, printing):
    max_time = int(max([x['end'] for x in [item for sublist in streams for item in sublist]]))+1
    start_cells = 'A2:A' + str(1+len(streams))
    instructions = 'r m' + str(max_time-1)
    turtles = [['!turtle(' + start_cells + ', ' + instructions + ', ' + str(speed) + ', 1)']]
    for stream in streams:
        cells = [""] * max_time
        for note in stream:
            start = int(note['start'])
            cells[start] = midi2str(note['note'])
            if note.has_key('velocity'):
                cells[start] += (' ' + str(round(float(note['velocity'])/127,2)))
            for rest_duration in range(1,int(note['length'])):
                cells[start+rest_duration] = '-'
        turtles.append(cells)
    cprint(str(len(turtles)) + ' x ' + str(max([len(stream) for stream in turtles])), printing)
    return turtles


# In[485]:


def midi_to_excello(file_name, method=1, logging=False, printing=True):
    # Fetch MIDI file
    mid = mido.MidiFile(file_name) 
    tempo = [m.dict()['tempo'] for m in mid.tracks[0] if m.dict().has_key('tempo')][0]
    ticks_per_beat = mid.ticks_per_beat
    # Extract the notes from as onset, note, offset, volume from messages
    all_notes = get_active_notes(mid)
    # Split into the streams as played by individual turtles
    streams = create_streams(all_notes)
    all_notes = [item for sublist in streams for item in sublist]
    cprint('Number of turtles: ' + str(len(streams)), printing)
    
    # No Compression
    if method == 0:
        cprint("No Compression", printing)
        difference_stat = 1
        ratio_int = 1
        for stream in streams:
            for note in stream:
                note['length'] = note['end'] - note['start']
    #Compression
    else:
        differences = [(y['start']-x['start']) for x, y in zip(all_notes[:-1], all_notes[1:])]
        lengths = [(x['end'] - x['start']) for x in [item for sublist in streams for item in sublist]]
        # Mins
        if method == 1:
            cprint("Min Compression", printing)
            difference_stat = min([x for x in differences if x > 1])
            length_stat = min([x for x in lengths if x > 1])
        # Modes 
        elif method == 2:
            cprint("Mode Compression", printing)
            difference_stat = max(set(differences), key=differences.count)
            length_stat = max(set(lengths), key=lengths.count)

        cprint('note difference stat: ' + str(difference_stat), printing)
        cprint('note length stat: ' + str(length_stat), printing)

        mode_ratio = (float(max(difference_stat, length_stat)) / min(difference_stat, length_stat))
        cprint('mode ratio: ' + str(mode_ratio), printing)
        ratio_int = int(mode_ratio)
        cprint('integer ratio: ' + str(ratio_int), printing)
#         ratio_correction = mode_ratio/ratio_int
#         cprint('ratio correction: ' + str(ratio_correction), printing)
    
        # Convert MIDI times to cell times
        rounding_base = 0.1
        for stream in streams:
            for note in stream:
                note['length'] = ((float(note['end']) - note['start'])/length_stat) 
                note['length'] = rounding_base * round(note['length']/rounding_base)
                note['start'] = round(rounding_base * round((float(note['start'])/difference_stat*ratio_int)/rounding_base))
                note['end'] = note['start'] + note['length']
            
    speed = int(round((float(60*10**6)/tempo) * ticks_per_beat * (float(ratio_int)/difference_stat)))
    cprint(speed, printing)
            
    csv_name = file_name[::-1].replace('/','_',file_name.count('/')-2)[::-1]
    csv_name = csv_name.replace('/midi','/csv/' + str(method)).replace('.mid','.csv')
    with open(csv_name, "wb") as f:
        writer = csv.writer(f)
        writer.writerows(streams_to_cells(streams, speed, printing))
    cprint("Written to " + csv_name, printing)
    
    if logging:
        cprint([csv_name, len(streams), int(max([x['end'] for x in [item for sublist in streams for item in sublist]]))], printing)
        return [csv_name, len(streams), int(max([x['end'] for x in [item for sublist in streams for item in sublist]]))]


# # Converting

# 0: No Compression<br>
# 1: Compression using Minimum difference<br>
# 2: Compression using Modal difference

# In[682]:


midi_to_excello('piano-midi/midi/debussy/DEB_CLAI.mid', 2)


# # Corpus Conversion

# In[692]:


datasets = ['piano-midi', 'bach', 'bach_chorales']


# In[693]:


def convert_corpus(corpus, method):
    midi_files = corpus + '/midi'
    files = []
    for r, _, f in os.walk(midi_files):
        for file in f:
            if '.mid' in file or '.MID' in file:
                files.append(os.path.join(r, file))
                
    if midi_files == 'bach/midi':
        files.remove('bach/midi/suites/airgstr4.mid')
        files = [ x for x in files if "wtcbki/" not in x ]
    
    log = []
    for f in files:
        log.append(midi_to_excello(f, method, logging=True, printing=False)) # This also writes the file to disk. 
    log.sort(key=lambda x: x[2], reverse=False)
    
    with open(midi_files.replace('/midi','/csv') + '/' + 'log' + str(method) + '.txt', mode="w") as outfile:
        outfile.write('%s\n'% len(log))
        for s in log:
            outfile.write("%s\n" % s)


# In[694]:


for corpus in datasets:
    for method in [0,1,2]:
        print(corpus, method)
        convert_corpus(corpus, method)


# # MIDI note name conversion test

# In[655]:


import audiolazy
for i in range(12,120):
    print(audiolazy.midi2str(i) == midi2str(i))

