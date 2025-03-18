import google.generativeai as genai

class APIHandler:
    def __init__(self, model_name='gemini-2.0-flash-lite'):
        self.model_name = model_name
        self.model = None

    def initialize_model(self, api_key):
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(self.model_name)

    def get_summary(self, text):
        if not self.model:
            return text

        try:
            prompt = f"""im going to give you the details of an event by the NSS Unit of Atlas SkillTech University. I need u to make a report for it which includes 2 paragraphs (dont mention paragraph 1, paragraph 2 etc, and it should be 200 words), each with a few lines about what took place, what was discussed, and mention The NSS Unit Of Atlas SkillTech University. here are the details: {text}"""
            response = self.model.generate_content(prompt)
            return response.text.strip()
        except Exception as e:
            raise Exception(f"Error in text summarization: {str(e)}")

    def get_takeaways(self, text, takeaway_prompt=None):
        if not self.model:
            return []
    
        try:
            if not takeaway_prompt:
                takeaway_prompt = """Based on the event description given, identify exactly five key problem-focused takeaways related to social service and substance abuse. Keep it a line or two each. Focus on the specific difficulties, dilemmas, and impacts discussed, rather than just general event outcomes. Format each takeaway with a title and description separated by a colon, without any asterisks (use numbers). Do not include the phrase 'Key Takeaways'. Example format:
                Takeaway Title: Takeaway"""
            
            prompt = takeaway_prompt + f"\nEvent Description: {text}"
            response = self.model.generate_content(prompt)
    
            takeaways = []
            for line in response.text.split('\n'):
                line = line.strip()
                #if line and not line.startswith(('•', '-', '*', '1.', '2.', '3.', '4.', '5.')):
                if line and not line.startswith(('•', '-', '*')):
                    line = line.replace('*', '')
                    takeaways.append(line.strip())
    
            return takeaways
        except Exception as e:
            raise Exception(f"Error generating takeaways: {str(e)}")
