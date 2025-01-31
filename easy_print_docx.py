import subprocess
import tempfile
import os

def print_docx(input_file_path, converter_executable='dependencies/docto.exe', print_executable='dependencies/SumatraPDF.exe'):
    """
    Converts a document (input_file_path) to the specified format (output_format) and sends it to the default printer.
    
    Parameters:
    - input_file_path: Path to the input document (e.g., DOCX, etc.)
    - converter_executable: Path to the DOCX-to-PDF converter executable (default: 'dependencies/docto.exe')
    - print_executable: Path to the executable that prints the document (default: 'dependencies/SumatraPDF.exe')
    
    Returns:
    - True if successful, False otherwise.
    """
    
    # Check if the input file exists
    if not os.path.isfile(input_file_path):
        print(f"Error: Input file does not exist: {input_file_path}")
        return False

    temp_output_path = None  # Initialize the variable to avoid potential reference before assignment

    try:
        output_format='pdf' 
        # Create a temporary file for the converted document
        with tempfile.NamedTemporaryFile(suffix=f".{output_format}", delete=False) as temp_output_file:
            temp_output_path = temp_output_file.name  # Store path before closing the file
        
        # Command to convert the document
        convert_command = [
            converter_executable,
            '-f', input_file_path,
            '-o', temp_output_path,
            '-t', f'wdFormat{output_format.upper()}'
        ]
        
        # Run the conversion command
        result = subprocess.run(convert_command, capture_output=True, text=True, check=True)
        print("Conversion successful!")
        print("Output:", result.stdout)
        
    except subprocess.CalledProcessError as e:
        print(f"Error converting file: {e.stderr}")
        if temp_output_path and os.path.isfile(temp_output_path):
            os.remove(temp_output_path)  # Cleanup on failure
        return False

    # Command to print the converted document
    print_command = [
        print_executable,
        '-print-to-default',
        temp_output_path
    ]
    
    try:
        # Send the print command
        subprocess.run(print_command, check=True)
        print("Print command sent successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error printing document: {e.stderr}")
        return False
    finally:
        # Cleanup the temporary file if it exists
        if temp_output_path and os.path.isfile(temp_output_path):
            os.remove(temp_output_path)
            print(f"Temporary file cleaned up: {temp_output_path}")

    return True


if __name__ == "__main__":
    # Example usage
    input_file = r"path_to_input_file.docx"  # Replace with the actual path
    if not print_docx(input_file):
        print("Conversion and print failed.")
