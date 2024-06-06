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
RAMP_STEP = 0.05  # Increase or decrease by this step

print("Listening...")

def set_spotify_volume_ramp(target_volume):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == "Spotify.exe":
            current_volume = volume.GetMasterVolume()
            while current_volume != target_volume:
                if current_volume < target_volume:
                    current_volume = min(current_volume + RAMP_STEP, target_volume)
                else:
                    current_volume = max(current_volume - RAMP_STEP, target_volume)
                volume.SetMasterVolume(current_volume, None)
                time.sleep(0.1)
            return True
    return False

try:
    while True:
        data = np.frombuffer(stream.read(CHUNK), dtype=np.int16)
        peak = np.abs(data).max()
        
        if peak > THRESHOLD:
            print("Loud noise detected, muting Spotify... Peak: " + str(peak))
            if set_spotify_volume_ramp(0.1):
                print("Spotify volume muted.")
            else:
                print("Failed to mute Spotify volume.")
            time.sleep(5)
            print("Restoring Spotify volume...")
            if set_spotify_volume_ramp(1.0):
                print("Spotify volume restored.")
            else:
                print("Failed to restore Spotify volume.")
            peak = 0
        
except KeyboardInterrupt:
    print("Exiting...")

finally:
    stream.stop_stream()
    stream.close()
    p.terminate()
