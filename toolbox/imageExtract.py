from transformers import Qwen2_5_VLForConditionalGeneration, BitsAndBytesConfig, AutoProcessor
from qwen_vl_utils import process_vision_info
import torch

def main(inputPath):
    min_pixels = 256*28*28
    max_pixels = 720*28*28
    modelName = "Qwen/Qwen2.5-VL-7B-Instruct"

    bnb_config = BitsAndBytesConfig(
            load_in_4bit=False,
            bnb_4bit_compute_dtype=torch.float16
    )

    model = Qwen2_5_VLForConditionalGeneration.from_pretrained(
        modelName, 
        torch_dtype="auto",
        device_map="auto",
        quantization_config=bnb_config,
    )

    processor = AutoProcessor.from_pretrained(
        modelName, min_pixels=min_pixels, max_pixels=max_pixels
    )


    messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "image": inputPath,
                },
                {"type": "text", "text": """
Extract all visible text from any provided image using OCR, and provide a detailed description of the image’s contents, layout, and relevant features.

- Always start by extracting all readable text within the image as accurately as possible.
- After text extraction, analyze and describe the overall image. Include details such as the setting, objects, people, actions, color scheme, and notable visual elements.
- For both text extraction and image description, reason step-by-step and ensure completeness before producing final outputs.
- If the text is incomplete, obscured, or not in English, note these issues and attempt partial extraction and description.
- If no text is present, clearly state this and focus solely on visual analysis.
- Present your output in a structured JSON format:

{
  "extracted_text": "[All extracted text as a single string. If none, return: 'No text detected.']",
  "image_description": "[A concise but thorough paragraph describing the visual elements and context.]"
}

Remember:  
First, extract and present all visible text (as a single string).  
Second, provide a paragraph describing the image’s content and context.  
Output must be a JSON object as above.
"""},
            ],
        }
    ]

    text = processor.apply_chat_template(
        messages, tokenize=False, add_generation_prompt=True
    )
    image_inputs, video_inputs = process_vision_info(messages)
    inputs = processor(
        text=[text],
        images=image_inputs,
        videos=video_inputs,
        padding=True,
        return_tensors="pt",
    )
    inputs = inputs.to("cuda")

    generated_ids = model.generate(**inputs, max_new_tokens=1024)
    generated_ids_trimmed = [
        out_ids[len(in_ids) :] for in_ids, out_ids in zip(inputs.input_ids, generated_ids)
    ]
    output_text = processor.batch_decode(
        generated_ids_trimmed, skip_special_tokens=True, clean_up_tokenization_spaces=False
    )
    
    return output_text[0]
