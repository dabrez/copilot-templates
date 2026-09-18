import json
import os
import sys
from pptx import Presentation
from llama_cpp import Llama
from huggingface_hub import hf_hub_download

def analyze_presentation_llm(input_path, output_path):
    """
    Analyzes a .pptx presentation for potential variables using a local LLM (Gemma 2B).
    
    :param input_path: Path to the source .pptx file.
    :param output_path: Path to save the generated JSON config file.
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # 1. Download/Load Model
    print("Loading LLM model (Gemma 2 2B)... This may take a moment on the first run.")
    try:
        model_path = hf_hub_download(
            repo_id="bartowski/gemma-2-2b-it-GGUF",
            filename="gemma-2-2b-it-Q4_K_M.gguf"
        )
    except Exception as e:
        raise RuntimeError(f"Failed to download model from Hugging Face: {e}")

    # 2. Initialize Model
    # n_ctx: context window size. 2048 should be enough for slide text.
    # n_threads: adjust based on CPU cores. Defaulting to 4.
    llm = Llama(model_path=model_path, n_ctx=2048, verbose=False)

    # 3. Extract Text
    prs = Presentation(input_path)
    all_text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    if paragraph.text.strip():
                        all_text.append(paragraph.text.strip())
            
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            for paragraph in cell.text_frame.paragraphs:
                                if paragraph.text.strip():
                                    all_text.append(paragraph.text.strip())

    full_text = "\n".join(all_text)

    # 4. Prompt the LLM
    prompt = f"""<start_of_turn>user
Identify all potential template variables in the following text from a PowerPoint presentation. 
Potential variables include:
- People's names
- Company/Organization names
- Dates
- Monetary values
- Specific locations
- Product names

Return the result STRICTLY as a JSON object where the keys are the text to find and the values are a generic label like "[ORG]", "[PERSON]", "[DATE]", etc.
Example output: {{"Acme Corp": "[ORG]", "John Doe": "[PERSON]"}}

Text:
---
{full_text}
---
JSON:<end_of_turn>
<start_of_turn>model
{{"""

    # We start the response with "{" to nudge the model toward JSON output.
    response = llm(prompt, max_tokens=512, stop=["<end_of_turn>"], temperature=0.1)
    json_text = "{" + response['choices'][0]['text']

    # 5. Parse and Save
    try:
        # Basic cleanup in case of trailing markdown or text
        if "}" in json_text:
            json_text = json_text[:json_text.rfind("}")+1]
        
        config = json.loads(json_text)
        
        with open(output_path, 'w') as f:
            json.dump(config, f, indent=4)
            
    except json.JSONDecodeError as e:
        print(f"Error parsing LLM output as JSON: {e}", file=sys.stderr)
        print(f"Raw output: {json_text}", file=sys.stderr)
        raise RuntimeError("LLM failed to generate a valid JSON configuration.")
