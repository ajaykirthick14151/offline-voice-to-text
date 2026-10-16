
import sounddevice as sd

print("Available audio devices:")
print(sd.query_devices())

print("\nRecording for 3 seconds. Speak now!")

audio = sd.rec(
    int(16000 * 3),
    samplerate=16000,
    channels=1,
    dtype="int16"
)

sd.wait()

print("Recording completed!")
print("Audio shape:", audio.shape)
