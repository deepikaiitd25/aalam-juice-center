class PolicyAgent:

    def validate(self, text: str) -> bool:
        # Simple placeholder policy
        if not text or len(text.strip()) == 0:
            return False
        return True
