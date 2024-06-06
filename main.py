import pyaudio
import numpy as np
import time
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume

CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 44100

p = pyaudio.PyAudio()

stream = p.open(format=FORMAT,
                channels=CHANNELS,
                rate=RATE,
                input=True,
                frames_per_buffer=CHUNK)

THRESHOLD = 15000

print("Listening...")

def set_spotify_volume(level):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == "Spotify.exe":
            volume.SetMasterVolume(level, None)
            return True
    return False

try:
    while True:
        data = np.frombuffer(stream.read(CHUNK), dtype=np.int16)
        peak = np.abs(data).max()
        
        if peak > THRESHOLD:
            print("Loud noise detected, muting Spotify..." + str(peak))
            set_spotify_volume(0.1)
            time.sleep(5)
            peak = 0
            print("Restoring Spotify volume..." + str(peak))
            set_spotify_volume(1.0)
        peak = 0

except KeyboardInterrupt:
    print("Exiting...")

finally:
    stream.stop_stream()
    stream.close()
    p.terminate()
