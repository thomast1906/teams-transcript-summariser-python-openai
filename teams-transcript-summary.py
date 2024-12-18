import os
import time
import logging
import json
from docx import Document
from dotenv import load_dotenv
from openai import AzureOpenAI

# Load environment variables from .env file
load_dotenv()

# Initialise logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialise Azure OpenAI Client
try:
    client = AzureOpenAI(
        api_key=os.getenv("AZURE_OPENAI_API_KEY"),  
        api_version="2024-02-01",
        azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT")
    )
except Exception as e:
    logging.error(f"Failed to initialize Azure OpenAI client: {e}")
    exit(1)

deployment_name = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME")

def read_and_chunk_document(doc_path, chunk_size=20000):
    """
    Reads a document and splits it into chunks of specified size.

    Args:
        doc_path (str): Path to the document file.
        chunk_size (int): Size of each chunk.

    Returns:
        list: List of text chunks.
    """
    try:
        doc = Document(doc_path)
        full_text = [para.text for para in doc.paragraphs]
        text = "\n".join(full_text)
        chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
        return chunks
    except Exception as e:
        logging.error(f"Error reading or chunking document: {e}")
        return []

def prompt_against_chunks(chunks, initial_prompt, client, deployment_name):
    """
    Prompts the model against each chunk of text.

    Args:
        chunks (list): List of text chunks.
        initial_prompt (list): Initial prompt messages from JSON.
        client (AzureOpenAI): Azure OpenAI client instance.
        deployment_name (str): Deployment name for the model.

    Returns:
        list: List of responses from the model.
    """
    responses = []
    for chunk in chunks:
        try:
            messages = initial_prompt + [{"role": "user", "content": chunk}]
            response = client.chat.completions.create(
                model=deployment_name,
                messages=messages,
                max_tokens=4096,
                temperature=0.5,
                top_p=0.5
            )
            if response.choices:
                responses.append(response.choices[0].message.content.strip())
            else:
                logging.warning(f"No choices returned for chunk: {chunk}")
        except Exception as e:
            logging.error(f"Error prompting against chunk: {e}")
    return responses

def main():
    # Record the start time
    start_time = time.time()

    # Check environment variables
    api_key = os.getenv("AZURE_OPENAI_API_KEY")
    endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
    deployment_name = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME")
    doc_path = os.getenv('DOC_PATH')

    if not api_key or not endpoint or not deployment_name or not doc_path:
        logging.error("One or more environment variables are missing.")
        return

    # Load prompts from prompts.json
    try:
        with open("prompts.json", "r") as file:
            prompts = json.load(file)
    except Exception as e:
        logging.error(f"Error loading prompts from JSON file: {e}")
        return

    initial_prompt_key = os.getenv("INITIAL_PROMPT_SUMMARY")
    final_prompt_key = os.getenv("FINAL_PROMPT_SUMMARY")
    initial_prompt = prompts[initial_prompt_key]
    final_prompt = prompts[final_prompt_key]

    # Read and chunk the document
    logging.info("Starting to chunk the document {}".format(doc_path))
    chunks = read_and_chunk_document(doc_path)

    if not chunks:
        logging.error("No chunks to process.")
        return

    # Get responses for each chunk
    responses = prompt_against_chunks(chunks, initial_prompt, client, deployment_name)

    # Output the responses to a file called responses.txt
    try:
        with open("responses.txt", "w") as file:
            for response in responses:
                file.write(response + "\n")
        logging.info("Responses written to responses.txt")
        logging.info("Finished chunking the document {}".format(doc_path))
    except Exception as e:
        logging.error(f"Error writing responses to file: {e}")
        return

    # Read the contents of responses.txt
    try:
        with open("responses.txt", "r") as file:
            responses_content = file.read()
    except Exception as e:
        logging.error(f"Error reading responses file: {e}")
        return

    # Prompt the model to summarise the entire content
    try:
        messages = final_prompt + [{"role": "user", "content": responses_content}]
        final_response = client.chat.completions.create(
            model=deployment_name,
            messages=messages,
            max_tokens=4096,
            temperature=0.5,
            top_p=0.5
        )
    except Exception as e:
        logging.error(f"Error generating final summary: {e}")
        return

    if final_response.choices:
        final_summary = final_response.choices[0].message.content.strip()
        logging.info("Final Summary has been created")

        # Save the final summary to a file called final_summary.txt
        try:
            with open("final_summary.txt", "w") as file:
                file.write(final_summary)
            logging.info("Final summary written to final_summary.txt")
        except Exception as e:
            logging.error(f"Error writing final summary to file: {e}")
    else:
        logging.warning("No summary returned.")

    # Record the end time
    end_time = time.time()  
    # Calculate the elapsed time
    elapsed_time = end_time - start_time  
    logging.info(f"Script execution time: {elapsed_time:.2f} seconds")

if __name__ == "__main__":
    main()