import openai
import docx2txt
from docx import Document
# Authenticate with OpenAI API
openai.api_key = "api key"
# Load English text from a Word document
english_text = docx2txt.process("example.docx")
# Translate English text to Chinese
translation = openai.Completion.create(
    engine="text-davinci-002",
    prompt=f"Translate the following text from English to Chinese:\n{english_text}",
    temperature=0.7,
    max_tokens=1024,
    n = 1,
    stop=None,
    frequency_penalty=0,
    presence_penalty=0
)
# Get the translated text
chinese_text = translation.choices[0].text
# Save the translated text as a Word document
document = Document()
document.add_paragraph(chinese_text)
document.save("output.docx")