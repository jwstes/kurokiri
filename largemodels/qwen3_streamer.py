import threading
import torch
from typing import Iterator, Optional
from transformers import (
    AutoModelForCausalLM,
    AutoTokenizer,
    BitsAndBytesConfig,
    TextIteratorStreamer,
)

# Default config
DEFAULT_MODEL_NAME = "Qwen/Qwen3-4B-Thinking-2507"
END_OF_THINK_TOKEN_ID = 151668

# Globals for reuse
_TOKENIZER = None
_MODEL = None


def init_llm(
    model_name: str = DEFAULT_MODEL_NAME,
    load_in_4bit: bool = False,
    compute_dtype: torch.dtype = torch.float16,
    device_map: str = "auto",
) -> None:
    """
    Load and cache the tokenizer/model once. Call this at startup.
    """
    global _TOKENIZER, _MODEL
    if _MODEL is not None and _TOKENIZER is not None:
        return

    bnb_config = BitsAndBytesConfig(
        load_in_4bit=load_in_4bit,
        bnb_4bit_compute_dtype=compute_dtype
    )

    _TOKENIZER = AutoTokenizer.from_pretrained(model_name)
    _MODEL = AutoModelForCausalLM.from_pretrained(
        model_name,
        torch_dtype="auto",
        device_map=device_map,
        quantization_config=bnb_config,
    )


def _ensure_initialized():
    if _MODEL is None or _TOKENIZER is None:
        raise RuntimeError("LLM not initialized. Call init_llm() first.")


def stream_prompt(
    prompt: str,
    max_new_tokens: int = 32768,
    hide_thinking: bool = True,
    end_of_think_token_id: int = END_OF_THINK_TOKEN_ID,
) -> Iterator[str]:
    """
    Start generation in a background thread and yield text chunks suitable for printing.

    - If hide_thinking=True, chunks before the special end-of-think marker are suppressed.
    - Otherwise, raw visible text is yielded as the model generates.

    Usage:
        for chunk in stream_prompt("Hello"):
            print(chunk, end="", flush=True)
    """
    _ensure_initialized()

    messages = [{"role": "user", "content": prompt}]
    text = _TOKENIZER.apply_chat_template(
        messages, tokenize=False, add_generation_prompt=True
    )
    model_inputs = _TOKENIZER([text], return_tensors="pt").to(_MODEL.device)

    # Keep special tokens when we need to detect the end-of-think marker
    streamer = TextIteratorStreamer(
        _TOKENIZER,
        skip_prompt=True,
        skip_special_tokens=not hide_thinking,
    )

    gen_kwargs = {
        "input_ids": model_inputs["input_ids"],
        "attention_mask": model_inputs.get("attention_mask"),
        "max_new_tokens": max_new_tokens,
        "streamer": streamer,
    }
    gen_kwargs = {k: v for k, v in gen_kwargs.items() if v is not None}

    thread = threading.Thread(target=_MODEL.generate, kwargs=gen_kwargs, daemon=True)
    thread.start()

    if hide_thinking:
        marker_text = _TOKENIZER.decode([end_of_think_token_id], skip_special_tokens=False)
        buffer = []
        saw_marker = False

        for chunk in streamer:
            if not saw_marker:
                buffer.append(chunk)
                joined = "".join(buffer)
                pos = joined.find(marker_text)
                if pos != -1:
                    saw_marker = True
                    after_marker = joined[pos + len(marker_text):]
                    if after_marker:
                        yield after_marker
                    buffer = None  # free buffer
            else:
                yield chunk

        # If the marker never appeared, emit whatever was buffered
        if not saw_marker:
            yield "".join(buffer)
    else:
        for chunk in streamer:
            yield chunk

    thread.join()


def generate_full(
    prompt: str,
    max_new_tokens: int = 32768,
    hide_thinking: bool = False,
    end_of_think_token_id: int = END_OF_THINK_TOKEN_ID,
) -> str:
    """
    Convenience function if you want the final concatenated text (no incremental printing).
    """
    chunks = []
    for c in stream_prompt(
        prompt,
        max_new_tokens=max_new_tokens,
        hide_thinking=hide_thinking,
        end_of_think_token_id=end_of_think_token_id,
    ):
        chunks.append(c)
    return "".join(chunks)