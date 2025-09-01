from transformers import WhisperProcessor, WhisperForConditionalGeneration
import librosa
import torch


def main(inputPath):
    device = "cuda:0" if torch.cuda.is_available() else "cpu"

    processor = WhisperProcessor.from_pretrained("openai/whisper-large-v2")
    model = WhisperForConditionalGeneration.from_pretrained("openai/whisper-large-v2").to(device)
    model.config.forced_decoder_ids = None

    audio_path = inputPath
    sample, sampling_rate = librosa.load(audio_path, sr=16000)


    input_features = processor(sample, sampling_rate=sampling_rate, return_tensors="pt").input_features
    input_features = input_features.to(device)

    predicted_ids = model.generate(input_features)


    transcription = processor.batch_decode(predicted_ids, skip_special_tokens=True)
    return transcription[0]