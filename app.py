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
You are an expert technical writer and business analyst. Create ONE consolidated, comprehensive functional document synthesizing information from the following Azure DevOps user stories. Aim for a coherent narrative, not just a list.

User Stories Data:
{consolidated_stories_info}

Structure the Consolidated Functional Document as follows:
1.  **Overall Objective / Epic Summary (if discernible)**
2.  **Key Features/Stories Covered** (Summarize main stories)
3.  **Detailed Functional Breakdown**: For each key area or synthesized feature:
    *   **Problem Statement & Impact**: Summarize core problems and impacts.
    *   **Functional Description**: Integrate details from story descriptions.
    *   **Acceptance Criteria**: Consolidate or list key acceptance criteria.
4.  **Cross-cutting Concerns / Shared Elements (if any)**

Maintain a professional tone.
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
st.title("Azure DevOps Functional Document Generator")
st.info(f"Using LLM: {llm_client.model_name if llm_client else 'N/A'}")

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
        # Add title (optional, can be more sophisticated)
        doc.add_heading('Combined Functional Document', level=1)
        
        # Split the markdown content by lines and add as paragraphs
        # This is a basic conversion. For full markdown to docx, a library like pandoc (via subprocess) or a dedicated Python markdown-to-docx converter would be better.
        # For now, we'll add line by line. Lines starting with #, ##, ===, --- could be styled.
        
        lines = st.session_state.generated_document.split('\n')
        for line in lines:
            if line.startswith('# '): # H1
                doc.add_heading(line[2:], level=1)
            elif line.startswith('## '): # H2
                doc.add_heading(line[3:], level=2)
            elif line.startswith('### '): # H3
                doc.add_heading(line[4:], level=3)
            elif line.startswith('====='): # Could be an H1 underline
                if doc.paragraphs and doc.paragraphs[-1].text: # Check if previous paragraph is not empty
                    # This is a simple heuristic, might not always be correct
                    # For now, we'll just skip these lines as they are part of markdown H1/H2 syntax
                    pass 
            elif line.startswith('-----'): # Could be an H2 underline or thematic break
                 # For now, we'll just skip these lines
                pass
            elif line.strip() == '---': # Thematic break
                doc.add_page_break() # Or some other separator
            elif line.strip().startswith('* ') or line.strip().startswith('- '): # Basic list item
                # python-docx can create bulleted lists, but requires more structure.
                # For simplicity, adding as a paragraph with the leading char.
                doc.add_paragraph(line)
            else:
                doc.add_paragraph(line)
        
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
