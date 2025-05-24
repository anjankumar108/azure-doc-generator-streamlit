import streamlit as st

# Set page config as the first Streamlit command
st.set_page_config(layout="wide")

import requests
import base64
import json
import re
import os
from dotenv import find_dotenv, load_dotenv
from langchain_openai import ChatOpenAI
from docx import Document # For creating .docx files
from docx.shared import Pt, RGBColor # For font size and color
from docx.enum.text import WD_ALIGN_PARAGRAPH # For aligning horizontal rule
from io import BytesIO # To handle document in memory for download

# --- Configuration Constants (mostly for Azure DevOps) ---
# These are still "hardcoded" in the script but are now clearly defined constants.
# For deployment (e.g., Streamlit Cloud), AZURE_DEVOPS_PAT_ENV will be read from secrets.
AZURE_DEVOPS_ORG_CONFIG = "landmarkgroup"
AZURE_DEVOPS_PROJECT_CONFIG = "S2"
AZURE_API_VERSION_CONFIG = "7.2-preview"
# Set to the field name provided by the user.
# IMPORTANT: This should be the *API Reference Name* from Azure DevOps,
# not just the display name. If "Current problem statement" is only the display name,
# fetching this field might fail or return N/A.
PROBLEM_IMPACT_FIELD_NAME_CONFIG = "Current problem statement"


# --- Environment and LLM Setup ---
try:
    dotenv_path = find_dotenv(raise_error_if_not_found=False)
    if dotenv_path:
        load_dotenv(dotenv_path)
    # No sidebar warning if .env not found, as sidebar is removed.
    # else:
    #     st.sidebar.warning("`.env` file not found. LITELLM_API_KEY should be set as an environment variable.")

    LITELLM_API_KEY = os.getenv("LITELLM_API_KEY")
    AZURE_DEVOPS_PAT_ENV = os.getenv("AZURE_DEVOPS_PAT") # Load PAT from environment

    if not LITELLM_API_KEY:
        st.error("LITELLM_API_KEY environment variable not set. Please set it in your .env file or Streamlit secrets. The app cannot function without it.")
        st.stop()
    # AZURE_DEVOPS_PAT_ENV can be None if not set, will be handled by UI fallback or error later

    # Cache the LLM client
    @st.cache_resource
    def get_llm_client_cached():
        if not LITELLM_API_KEY:
            st.error("LITELLM_API_KEY is missing, cannot initialize LLM client.")
            return None
        try:
            client = ChatOpenAI(
                openai_api_base="https://lmlitellm.landmarkgroup.com", # Ensure this is correct
                default_headers={
                    "Authorization": f"Bearer {LITELLM_API_KEY}",
                    "Content-Type": "application/json"
                },
                model="gemini-2.5-pro",
                api_key=LITELLM_API_KEY,
                temperature=0.7
            )
            return client
        except Exception as e:
            st.error(f"Failed to initialize LLM client: {e}")
            return None

    llm_client = get_llm_client_cached() # Use the cached version
    if not llm_client:
        st.stop()

except Exception as e:
    st.error(f"Error during initial setup: {e}")
    st.stop()


# --- Core Logic Functions (now use hardcoded config) ---

def extract_work_item_id(input_string):
    match = re.search(r'/(\d+)/?$', input_string)
    if match:
        return match.group(1)
    elif input_string.isdigit():
        return input_string
    return None

# Updated to accept PAT as an argument
def get_azure_devops_story_details(work_item_id, pat_to_use, progress_bar_slot):
    if not pat_to_use:
        st.error("Azure DevOps PAT was not provided or found.")
        return None

    url = f"https://dev.azure.com/{AZURE_DEVOPS_ORG_CONFIG}/{AZURE_DEVOPS_PROJECT_CONFIG}/_apis/wit/workitems/{work_item_id}?$expand=all&api-version={AZURE_API_VERSION_CONFIG}"
    
    try:
        credentials = f":{pat_to_use}"
        encoded_credentials = base64.b64encode(credentials.encode()).decode()
        headers = {"Authorization": f"Basic {encoded_credentials}", "Content-Type": "application/json"}
        
        progress_bar_slot.info(f"Fetching details for work item ID: {work_item_id}...")
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()
        
        details = {
            "id": data.get("id"),
            "title": data.get("fields", {}).get("System.Title", "N/A"),
            "description": data.get("fields", {}).get("System.Description", "N/A"),
            "acceptance_criteria": data.get("fields", {}).get("Microsoft.VSTS.Common.AcceptanceCriteria", "N/A"),
            "problem_impact": data.get("fields", {}).get(PROBLEM_IMPACT_FIELD_NAME_CONFIG, "N/A") # Uses hardcoded field name
        }
        progress_bar_slot.success(f"Successfully fetched details for ID: {work_item_id}")
        return details
        
    except requests.exceptions.HTTPError as http_err:
        st.error(f"HTTP error for ID {work_item_id}: {http_err}. Response: {response.text}")
    except requests.exceptions.RequestException as req_err:
        st.error(f"Request error for ID {work_item_id}: {req_err}")
    except json.JSONDecodeError:
        st.error(f"Failed to decode JSON for ID {work_item_id}. Response: {response.text}")
    except Exception as e:
        st.error(f"An unexpected error occurred for ID {work_item_id}: {e}")
    return None

def generate_combined_functional_document(all_stories_data, llm_instance, progress_bar_slot):
    if not all_stories_data:
        return "No story data provided to generate a document."

    def clean_text(text_content):
        if text_content is None: return "N/A"
        text_content = re.sub('<[^>]+>', '', text_content)
        html_entities = {"&nbsp;": " ", "<": "<", ">": ">", "&": "&"}
        for entity, char in html_entities.items():
            text_content = text_content.replace(entity, char)
        return text_content.strip()

    stories_summary_parts = []
    fallback_content_parts = ["Combined Functional Document (LLM Failed - Basic Fallback)\n==========================================================\n"]

    for i, details in enumerate(all_stories_data):
        story_id = details.get('id', 'N/A')
        story_title = details.get('title', 'N/A')
        description_text = clean_text(details.get('description'))
        acceptance_criteria_text = clean_text(details.get('acceptance_criteria'))
        problem_impact_text = clean_text(details.get('problem_impact'))

        stories_summary_parts.append(f"---\nStory {i+1} (ID: {story_id})\nTitle: \"{story_title}\"\nProblem/Impact: {problem_impact_text}\nDescription: {description_text}\nAcceptance Criteria: {acceptance_criteria_text}\n---")
        fallback_content_parts.extend([f"\n--- Details for Story ID: {story_id} ---\n", f"Title: {story_title}\n", f"Problem/Impact: {problem_impact_text}\n", f"Description: {description_text}\n", f"Acceptance Criteria: {acceptance_criteria_text}\n"])

    consolidated_stories_info = "\n".join(stories_summary_parts)
    combined_fallback_text = "".join(fallback_content_parts)

    prompt = f"""
You are an expert technical writer and business analyst. Your task is to create ONE consolidated, comprehensive functional document by synthesizing information from the provided Azure DevOps user stories.

**Output Requirements:**
1.  **Document Title:** Begin your response *directly* with a concise and descriptive title for the functional document. Do not include any introductory phrases like "Okay, here is..." or "Consolidated Functional Document:". For example, a good title might be "Enhanced Short-Pick Handling in WMS".
2.  **Document Body:** Following the title, structure the document as follows:
    *   **1. Overall Objective / Epic Summary (if discernible from the stories)**
    *   **2. Key Functionalities Addressed:**
        *   Describe the core functionalities or features covered by this document in a narrative style.
        *   Integrate the essence of the user stories provided.
        *   Avoid directly mentioning "user story IDs" or presenting it as a list of Azure DevOps items. Instead, explain what enhancements or capabilities are being introduced. For example, instead of "Story 123: Implement X", say "This document outlines the implementation of X, which allows users to...".
    *   **3. Detailed Functional Breakdown**: For each key area or synthesized feature:
        *   **Context and Rationale**: Summarize the background, core problems, and their business impacts that necessitate this functionality.
        *   **Functional Description**: Provide an integrated description of how the functionality works, drawing from the user story details.
        *   **Verification Criteria**: Consolidate or list key criteria that will be used to confirm the functionality is implemented correctly, in a clear, testable format.
    *   **4. Cross-cutting Concerns / Shared Elements (if any relevant)**

**Important Style Guidelines:**
*   Maintain a professional and formal tone throughout the document.
*   The goal is to produce a document that reads like a standard, well-written functional specification.
*   **Avoid any meta-commentary about the document itself or internal team communications.** For example, do not include sentences like "Technical documentation detailing backend changes... must be developed and shared with the Quality Assurance (QA) team..." or "Comprehensive user documentation will be required...". Focus solely on describing the functionality.

User Stories Data:
{consolidated_stories_info}
    """
    try:
        progress_bar_slot.info("Invoking LLM to generate combined document... (this may take a moment)")
        response = llm_instance.invoke(prompt)
        llm_generated_content = response.content if hasattr(response, 'content') else str(response)
        progress_bar_slot.success("LLM processing complete.")
        return llm_generated_content
    except Exception as e:
        st.error(f"Error invoking LLM: {e}")
        return combined_fallback_text

# --- Streamlit App UI (Simplified, PAT from env or UI fallback) ---
st.title("Functional Document Generator")

# Get PAT: Try from environment (for deployed app), fallback to UI input for local dev
azure_pat_from_env = os.getenv("AZURE_DEVOPS_PAT")
pat_input = azure_pat_from_env # Default to env var

if not azure_pat_from_env: # If not in env, show input field
    st.warning("Azure DevOps PAT not found in environment variables. Please enter it below for this session.")
    pat_input = st.text_input("Azure DevOps Personal Access Token (PAT)", type="password", help="Your Azure DevOps PAT with read access to work items.")

# Main Area for Input and Output
st.header("Input Work Items")
story_ids_input = st.text_area(
    f"Enter Azure DevOps story links or IDs for {AZURE_DEVOPS_ORG_CONFIG}/{AZURE_DEVOPS_PROJECT_CONFIG} (comma-separated)", 
    height=100, 
    placeholder=f"e.g., 12345, https://dev.azure.com/{AZURE_DEVOPS_ORG_CONFIG}/{AZURE_DEVOPS_PROJECT_CONFIG}/_workitems/edit/67890"
)

if 'generated_document' not in st.session_state:
    st.session_state.generated_document = ""
if 'processed_ids_for_run' not in st.session_state: # To store IDs for download filename
    st.session_state.processed_ids_for_run = []


status_placeholder = st.empty()

if st.button("Generate Document", type="primary", use_container_width=True):
    current_pat_to_use = pat_input # This will be from env or UI input
    
    if not current_pat_to_use:
        st.error("Azure DevOps PAT is required. Please set the AZURE_DEVOPS_PAT environment variable or enter it above.")
    elif not story_ids_input:
        st.error("Please enter at least one story link or ID.")
    elif not PROBLEM_IMPACT_FIELD_NAME_CONFIG:
         st.error("The 'Problem/Impact Field API Name' (PROBLEM_IMPACT_FIELD_NAME_CONFIG) is empty in app.py. Please contact the app administrator.")
    else:
        st.session_state.generated_document = ""
        st.session_state.processed_ids_for_run = []
        status_placeholder.info("Starting document generation process...")
        
        input_items = [item.strip() for item in story_ids_input.split(',') if item.strip()]
        all_stories_data = []
        has_errors = False

        if not input_items:
            st.error("No valid story links or IDs provided after parsing.")
        else:
            for item_val in input_items:
                work_item_id = extract_work_item_id(item_val)
                if work_item_id:
                    story_data = get_azure_devops_story_details(
                        work_item_id, current_pat_to_use, status_placeholder
                    )
                    if story_data:
                        all_stories_data.append(story_data)
                        st.session_state.processed_ids_for_run.append(str(work_item_id))
                    else:
                        has_errors = True 
                else:
                    status_placeholder.warning(f"Skipping invalid input: '{item_val}'")
            
            if all_stories_data and not has_errors:
                status_placeholder.info(f"All data fetched for IDs: {', '.join(st.session_state.processed_ids_for_run)}. Generating combined document...")
                combined_doc = generate_combined_functional_document(all_stories_data, llm_client, status_placeholder)
                st.session_state.generated_document = combined_doc
                if "LLM Failed" not in combined_doc:
                     status_placeholder.success("Document generation complete!")
            elif not all_stories_data:
                 status_placeholder.error("No valid story data could be fetched. Cannot generate document.")
            else:
                 status_placeholder.warning("Document generation skipped or incomplete due to errors in fetching some work items. Please review messages above.")

if st.session_state.generated_document:
    st.header("Generated Functional Document")
    st.markdown(st.session_state.generated_document)
    
    # Create a unique filename for download
    # Use st.session_state.processed_ids_for_run which is now correctly populated
    download_filename_suffix = "_".join(st.session_state.processed_ids_for_run) if st.session_state.processed_ids_for_run and len(st.session_state.processed_ids_for_run) < 5 else "multiple_stories"
    
    # --- Create DOCX for download ---
    try:
        doc = Document()

        # Define and apply styles
        # Normal text style
        style_normal = doc.styles['Normal']
        font_normal = style_normal.font
        font_normal.size = Pt(13)

        # Define a blue color
        blue_color = RGBColor(0x1E, 0x88, 0xE5) # A pleasant shade of blue

        # Heading 1 style
        style_h1 = doc.styles['Heading 1']
        font_h1 = style_h1.font
        font_h1.size = Pt(20)
        font_h1.color.rgb = blue_color
        font_h1.bold = True 
        para_format_h1 = style_h1.paragraph_format
        para_format_h1.space_after = Pt(12) # Add 12pt space after H1

        # Heading 2 style
        style_h2 = doc.styles['Heading 2']
        font_h2 = style_h2.font
        font_h2.size = Pt(16)
        font_h2.color.rgb = blue_color
        font_h2.bold = True
        para_format_h2 = style_h2.paragraph_format
        para_format_h2.space_after = Pt(8) # Add 8pt space after H2

        # Heading 3 style
        style_h3 = doc.styles['Heading 3']
        font_h3 = style_h3.font
        font_h3.size = Pt(14)
        font_h3.color.rgb = blue_color
        font_h3.bold = True
        para_format_h3 = style_h3.paragraph_format
        para_format_h3.space_after = Pt(6) # Add 6pt space after H3

        full_markdown_content = st.session_state.generated_document

        # Pre-processing for the entire content (e.g., em space)
        full_markdown_content = full_markdown_content.replace('\u2003', ' ') 
        full_markdown_content = full_markdown_content.strip()

        # Extract title (first line) and the rest of the content
        content_lines = full_markdown_content.split('\n', 1)
        doc_title_raw = content_lines[0].strip()
        markdown_body = content_lines[1].strip() if len(content_lines) > 1 else ""

        # Clean the extracted title (remove potential markdown heading markers)
        doc_title_cleaned = re.sub(r'^[#\s]*', '', doc_title_raw)
        
        # Further pre-processing for the body
        markdown_body = re.sub(r'(\r\n|\r|\n){3,}', '\n\n', markdown_body) # Consolidate >2 newlines
        markdown_body = markdown_body.strip()

        # Helper function to add text with inline formatting to a paragraph object
        def add_runs_to_paragraph(paragraph, text_content):
            # Regex to split by bold/italic markers, non-greedy
            # Handles **bold**, *italic*, __bold__, _italic_
            segments = re.split(r'(\*\*(?:(?!\*\*).)*\*\*|\*(?:(?!\*).)*\*|__(?:(?!__).)*__|_(?:(?!_).)*_)', text_content)
            
            for segment in segments:
                if not segment: # Skip empty strings that can result from split
                    continue
                
                run = paragraph.add_run()
                if (segment.startswith('**') and segment.endswith('**')) or \
                   (segment.startswith('__') and segment.endswith('__')): # Bold
                    run.text = segment[2:-2]
                    run.bold = True
                elif (segment.startswith('*') and segment.endswith('*')) or \
                     (segment.startswith('_') and segment.endswith('_')): # Italic
                    run.text = segment[1:-1]
                    run.italic = True
                else: # Regular text
                    run.text = segment

        # Add the extracted and cleaned title to the document
        if doc_title_cleaned:
            title_heading = doc.add_heading(level=1) 
            add_runs_to_paragraph(title_heading, doc_title_cleaned)

        # Process line by line from markdown_body
        lines = markdown_body.split('\n')
        paragraph_buffer = [] # To collect lines for a single paragraph

        # Note: The main document title (level 0 or 1) is now handled above.
        # The loop below will handle subsequent headings (H1, H2, H3 from markdown as H1,H2,H3 in docx)

        for i, line_raw in enumerate(lines):
            line = line_raw.strip()

            # Handle paragraph breaks (empty line)
            if not line:
                if paragraph_buffer:
                    p_text = " ".join(paragraph_buffer) # Join lines with space
                    add_runs_to_paragraph(doc.add_paragraph(), p_text)
                    paragraph_buffer = []
                continue 

            # Headings
            if line.startswith('# '):
                if paragraph_buffer: add_runs_to_paragraph(doc.add_paragraph(), " ".join(paragraph_buffer)); paragraph_buffer = []
                heading_text = line[2:].strip()
                h = doc.add_heading(level=1)
                add_runs_to_paragraph(h, heading_text)
            elif line.startswith('## '):
                if paragraph_buffer: add_runs_to_paragraph(doc.add_paragraph(), " ".join(paragraph_buffer)); paragraph_buffer = []
                heading_text = line[3:].strip()
                h = doc.add_heading(level=2)
                add_runs_to_paragraph(h, heading_text)
            elif line.startswith('### '):
                if paragraph_buffer: add_runs_to_paragraph(doc.add_paragraph(), " ".join(paragraph_buffer)); paragraph_buffer = []
                heading_text = line[4:].strip()
                h = doc.add_heading(level=3)
                add_runs_to_paragraph(h, heading_text)
            # Thematic break (---, ***, ___)
            elif line == '---' or line == '***' or line == '___':
                if paragraph_buffer: add_runs_to_paragraph(doc.add_paragraph(), " ".join(paragraph_buffer)); paragraph_buffer = []
                if line == '---': # '---' is treated as a page break
                     doc.add_page_break()
                else: # For ***, ___ draw a visual horizontal line
                     hr_p = doc.add_paragraph()
                     hr_p.add_run("_________________________________________")
                     hr_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Unordered List items (*, -, +)
            elif line.startswith('* ') or line.startswith('- ') or line.startswith('+ '):
                if paragraph_buffer: add_runs_to_paragraph(doc.add_paragraph(), " ".join(paragraph_buffer)); paragraph_buffer = []
                item_text = line[2:].strip()
                p = doc.add_paragraph(style='ListBullet')
                add_runs_to_paragraph(p, item_text)
            # Ordered List items (1., 2.)
            elif re.match(r'^\d+\.\s', line):
                if paragraph_buffer: add_runs_to_paragraph(doc.add_paragraph(), " ".join(paragraph_buffer)); paragraph_buffer = []
                item_text = re.sub(r'^\d+\.\s', '', line).strip()
                p = doc.add_paragraph(style='ListNumber')
                add_runs_to_paragraph(p, item_text)
            # Regular text line, add to buffer
            else:
                paragraph_buffer.append(line_raw) # Keep original spacing for paragraph internal lines

        # Add any remaining text in the buffer
        if paragraph_buffer:
            p_text = "\n".join(paragraph_buffer) # Join lines with newline for multi-line paragraphs
            add_runs_to_paragraph(doc.add_paragraph(), p_text)
        
        # Save to a BytesIO object
        bio = BytesIO()
        doc.save(bio)
        bio.seek(0)
        
        st.download_button(
            label="Download Document as Word File (.docx)",
            data=bio,
            file_name=f"functional_doc_combined_{download_filename_suffix}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Error creating .docx file: {e}")
        # Fallback to text download if docx creation fails
        st.download_button(
            label="Download Document as Text File (DOCX failed)",
            data=st.session_state.generated_document,
            file_name=f"functional_doc_combined_{download_filename_suffix}.txt",
            mime="text/plain"
        )

st.markdown("---")
