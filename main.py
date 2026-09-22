import asyncio
import json
import openpyxl

from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter 

from config import get_settings, get_async_openai
from glossary import fetch_glossary, match_terms, format_terms
from prompt import build_lqa_prompt


async def _complete_once(settings, prompt: str, model_name: str) -> str:
    """Sends a single completion request using a specific model name."""
    client = get_async_openai(settings)
    
    # 1. Print right BEFORE it goes to the network
    print("\n--- [START] Sending request to OpenRouter... ---")
    
    response = await client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=settings.temperature,
        response_format={"type": "json_object"},
        extra_body={"reasoning": {"enabled": False}},
    )
    
    # 2. Print right AFTER it comes back
    print("+++ [SUCCESS] Received response from OpenRouter! +++")
    
    if not response.choices:
        raise ValueError("Model returned no choices.")
    
    content = response.choices[0].message.content
    if not content:
        raise ValueError("Model returned empty message content.")
        
    return content


async def evaluate_row(
    item: dict,
    settings,
    source_lang: str,
    target_lang: str,
    semaphore: asyncio.Semaphore,
    cancel_flag: list = None  # <-- Added flag
) -> dict:
    """Evaluates a single text pair concurrently with fallback models and retries."""
    
    # 1. Check if stopped before waiting in line
    if cancel_flag and cancel_flag[0]:
        return item
        
    async with semaphore:
        # 2. Check if stopped right before building prompt and making API call
        if cancel_flag and cancel_flag[0]:
            return item
            
        prompt = build_lqa_prompt(
            source_text=item["source"],
            target_text=item["target"],
            source_lang=source_lang,
            target_lang=target_lang,
            formatted_terms=item["formatted_terms"],
        )
        
        models_list = settings.openrouter_models
        
        # Retry loop with Fallback Models
        for attempt in range(settings.max_retries):
            # 3. Check if stopped during a retry wait
            if cancel_flag and cancel_flag[0]:
                return item
                
            # Pick the model based on the attempt number. 
            # If attempts > available models, it just keeps trying the last model in the list.
            current_model = models_list[min(attempt, len(models_list) - 1)]
            
            try:
                raw_response = await _complete_once(settings, prompt, current_model)
                result = json.loads(raw_response)
                
                item["ai_status"] = result.get("issue", "")
                item["ai_severity"] = result.get("severity","")
                item["ai_feedback"] = result.get("reason", "")
                item["ai_suggestion"] = result.get("suggestion", "")
                
                return item
                
            except Exception as e:
                if attempt < settings.max_retries - 1:
                    # Exponential backoff before trying the next model
                    await asyncio.sleep(2 ** attempt)
                else:
                    item["ai_status"] = "系统错误"
                    item["ai_severity"] = "系统错误"
                    item["ai_feedback"] = f"重试 {settings.max_retries} 次后失败。最后使用的模型: {current_model}。错误: {str(e)}"
                    item["ai_suggestion"] = ""
                    
        return item


async def process_batch(items: list[dict], source_lang: str, target_lang: str, cancel_flag: list = None) -> list[dict]:
    """Processes all items concurrently with an active Watcher to kill in-flight network requests."""
    settings = get_settings()
    semaphore = asyncio.Semaphore(settings.batch_size)
    
    # --- NEW: Active Cancellation Watcher ---
    async def cancel_watcher():
        while True:
            if cancel_flag and cancel_flag[0]:
                print("\nDEBUG: Watcher saw Stop flag! Nuking all 100 in-flight network requests...")
                # Loop through every active task and forcefully drop the network connection
                for task in asyncio.all_tasks():
                    if task is not asyncio.current_task():
                        task.cancel()
                break
            await asyncio.sleep(0.2) # Check the Stop button 5 times a second
            
    # Start the watcher in the background
    watcher_task = asyncio.create_task(cancel_watcher())
    
    # Queue up all the rows
    tasks = [
        asyncio.create_task(evaluate_row(item, settings, source_lang, target_lang, semaphore, cancel_flag))
        for item in items
    ]
    
    try:
        # return_exceptions=True ensures that when the Watcher violently kills the tasks,
        # the program doesn't crash, but safely collects the cancelled pieces.
        results = await asyncio.gather(*tasks, return_exceptions=True)
        
        # Only return rows that successfully finished (filters out the cancelled ones)
        valid_results = [r for r in results if isinstance(r, dict)]
        return valid_results
    finally:
        # Clean up the watcher when the batch finishes successfully
        watcher_task.cancel()


def run_lqa(
    wb: openpyxl.Workbook,
    source_lang: str,
    target_lang: str,
    game_id: str | None,
    selected_sheet: str,
    use_glossary: bool = False,
    cancel_flag: list = None  # <-- Added flag
) -> openpyxl.Workbook:
    """
    Main orchestration pipeline called by Streamlit (app.py):
    """
    settings = get_settings()
    ws = wb[selected_sheet]
    
    # Clean headers exactly as we did in app.py to ensure perfect matching
    raw_headers = [cell.value for cell in ws[1]]
    headers = [
        str(h).strip() if h is not None and str(h).strip() != "" else f"Unnamed_{i+1}"
        for i, h in enumerate(raw_headers)
    ]
    
    try:
        src_idx = headers.index(source_lang) + 1  # openpyxl is 1-indexed
    except ValueError:
        raise ValueError(f"Source language column '{source_lang}' not found in worksheet headers.")

    try:
        tgt_idx = headers.index(target_lang) + 1
    except ValueError:
        raise ValueError(f"Target language column '{target_lang}' not found in worksheet headers.")

    # Step 1: Fetch glossary terms from Supabase if toggled
    game_terms = []
    if use_glossary and game_id:
        try:
            game_terms = fetch_glossary(settings, game_id, target_lang)
        except Exception as e:
            print(f"Glossary fetch warning: {e}")

    # Step 2: Build dataset from sheet rows
    items = []
    for row_idx in range(2, ws.max_row + 1):
        source_cell = ws.cell(row=row_idx, column=src_idx)
        source_text = str(source_cell.value or "").strip()
        
        if not source_text:
            continue

        target_cell = ws.cell(row=row_idx, column=tgt_idx)
        target_text = str(target_cell.value or "").strip()

        matched_terms = match_terms(game_terms, source_text) if game_terms else []

        items.append({
            "row": row_idx,
            "source": source_text,
            "target": target_text,
            "terms": matched_terms,
            "formatted_terms": format_terms(matched_terms),
        })

    if not items:
        return wb

    # Step 3: Run async LLM processing loop
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        # Pass the cancel_flag into process_batch
        processed_items = loop.run_until_complete(
            process_batch(items, source_lang, target_lang, cancel_flag)
        )
    except BaseException as e:
        # This catches Streamlit's internal Stop/Rerun signal!
        print("\nDEBUG: Stop signal received! Cancelling all LLM tasks...")
        
        # Force-cancel every single background task instantly
        for task in asyncio.all_tasks(loop):
            task.cancel()
            
        # Give the loop a split second to process the cancellations
        loop.run_until_complete(asyncio.sleep(0))
        
        # Pass the stop signal back to Streamlit so it can reset the UI
        raise e
    finally:
        loop.close()

    # Step 4: Write output back to openpyxl sheet (overwrite if exists, append if new)
    headers = [cell.value for cell in ws[1]]
    output_cols = ["严重级别","错误类型", "原因", "建议翻译"]
    col_indices = {}

    for col_name in output_cols:
        if col_name in headers:
            col_indices[col_name] = headers.index(col_name) + 1
        else:
            new_col_idx = ws.max_column + 1
            ws.cell(row=1, column=new_col_idx, value=col_name)
            headers.append(col_name)
            col_indices[col_name] = new_col_idx

    # --- Formatting Setup ---
    # 1. Set column widths
    widths = {"严重级别":10,"错误类型": 18, "原因": 45, "建议翻译": 35}
    for col_name, col_idx in col_indices.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = widths.get(col_name, 20)

    # 2. Define text wrap and vertical alignment
    wrap_align = Alignment(wrap_text=True, vertical="top")

    # --- Write Data and Apply Formatting ---
    for item in processed_items:
        r = item["row"]
        
        cell_status = ws.cell(row=r, column=col_indices["严重级别"], value=item.get("ai_severity"))
        cell_status.alignment = wrap_align

        cell_status = ws.cell(row=r, column=col_indices["错误类型"], value=item.get("ai_status"))
        cell_status.alignment = wrap_align
        
        cell_feedback = ws.cell(row=r, column=col_indices["原因"], value=item.get("ai_feedback"))
        cell_feedback.alignment = wrap_align
        
        cell_suggestion = ws.cell(row=r, column=col_indices["建议翻译"], value=item.get("ai_suggestion"))
        cell_suggestion.alignment = wrap_align

    return wb