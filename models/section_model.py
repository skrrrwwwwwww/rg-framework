class SectionModel:
    """Хранит структуру разделов и их содержимое"""
    def __init__(self):
        self.sections = {}  # {name: {"text": "", "subsections": []}}

    def add_section(self, name):
        self.sections[name] = {"text": "", "subsections": []}

    def add_subsection(self, parent, name):
        if parent not in self.sections:
            self.sections[parent] = {"text": "", "subsections": []}
        self.sections[parent]["subsections"].append({"name": name, "text": ""})

    def get_section_text(self, name, parent=None):
        if parent:
            for sub in self.sections[parent]["subsections"]:
                if sub["name"] == name:
                    return sub["text"]
            return ""
        return self.sections.get(name, {}).get("text", "")

    def get_all_sections(self):
        """Возвращает все разделы (для генерации)"""
        return self.sections

    def set_section_text(self, name, text, parent=None):
        if parent:
            for sub in self.sections[parent]["subsections"]:
                if sub["name"] == name:
                    sub["text"] = text
                    return
        else:
            if name not in self.sections:
                self.sections[name] = {"text": "", "subsections": []}
            self.sections[name]["text"] = text

    def remove_section(self, name):
        if name in self.sections:
            del self.sections[name]

    def remove_subsection(self, parent, name):
        if parent in self.sections:
            self.sections[parent]["subsections"] = [
                s for s in self.sections[parent]["subsections"] if s["name"] != name
            ]