import whisperx
import sys
import subprocess
import tempfile

# Entrée : vidéo MP4
video_path = sys.argv[1]

# Étape 1 : Transcription + diarisation
model = whisperx.load_model("large-v2", device="cuda" if whisperx.torch.cuda.is_available() else "cpu")
audio = whisperx.load_audio(video_path)
result = model.transcribe(audio)

# Étape 2 : Diarisation
diaraize_model = whisperx.DiarizationPipeline(use_auth_token=None)
segments = diarize_model(audio, min_speakers=2, max_speakers=6)

# Combine transcription et diarisation
final_result = whisperx.align_transcription(result["segments"], segments, model.sample_rate)
full_text = ""
for seg in final_result["segments"]:
    full_text += f"[{seg['speaker']}] {seg['text'].strip()}\n"

# Étape 3 : Résumé via Ollama
PROMPT = f"""
Tu es un assistant administratif. Voici une transcription brute d'une audition parlementaire.
Génère un compte rendu structuré destiné à l’administration, avec les sections suivantes :
- Liste des participants
- Résumé par thème
- Citations clés
- Conclusion

Texte source :
{full_text}
"""

process = subprocess.Popen(
    ["ollama", "run", "mistral"],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True
)
stdout, stderr = process.communicate(PROMPT)

# Export du compte rendu
with open("resume.txt", "w") as out:
    out.write(stdout)
