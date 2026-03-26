import json
import logging
import inspect
from typing import Any

from a2a.server.agent_execution import AgentExecutor
from a2a.server.agent_execution.context import RequestContext
from a2a.server.events.event_queue import EventQueue
from a2a.server.tasks import TaskUpdater
from a2a.types import (
    AgentCard,
    TaskState,
    TextPart,
    UnsupportedOperationError,
)
from a2a.utils.errors import ServerError
from openai import AsyncOpenAI

logger = logging.getLogger(__name__)


class OpenAIAgentExecutor(AgentExecutor):
    """
    A generic Executor that bridges the A2A protocol with OpenAI-compatible LLMs (like Gemini).
    It handles dynamic tool calling and structured data extraction.
    """

    def __init__(
        self,
        card: AgentCard,
        tools: dict[str, Any],
        api_key: str,
        system_prompt: str,
        base_url: str | None = None,
        model: str = "gemini-2.5-flash",
    ):
        self._card = card
        self.tools = tools
        self.client = AsyncOpenAI(
            api_key=api_key,
            base_url=base_url,
        )
        self.model = model
        self.system_prompt = system_prompt

    async def _process_request(
        self,
        message_text: str,
        context: RequestContext,
        task_updater: TaskUpdater,
    ) -> None:
        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": message_text},
        ]

        # Convert your Python tools into OpenAI function-calling format
        openai_tools = []
        for tool_name, tool_instance in self.tools.items():
            if hasattr(tool_instance, tool_name):
                func = getattr(tool_instance, tool_name)
                # Enhanced schema extraction for structured data (lists/dicts)
                schema = self._extract_function_schema(func)
                openai_tools.append({"type": "function", "function": schema})

        try:
            response = await self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                tools=openai_tools if openai_tools else None,
                tool_choice="auto" if openai_tools else None,
                temperature=0.1,
            )

            message = response.choices[0].message

            # Handle Tool Calls (e.g., Calling generate_docx)
            if message.tool_calls:
                for tool_call in message.tool_calls:
                    function_name = tool_call.function.name
                    function_args = json.loads(tool_call.function.arguments)

                    logger.info(f"🛠️ Executing tool: {function_name}")

                    if function_name in self.tools:
                        tool_provider = self.tools[function_name]
                        method = getattr(tool_provider, function_name)

                        # Execute the tool (generate_docx / generate_pptx)
                        result = method(**function_args)

                        # Format the result for the A2A Frontend
                        if hasattr(result, "model_dump_json"):
                            result_text = result.model_dump_json()
                        else:
                            result_text = json.dumps(result) if isinstance(
                                result, dict) else str(result)

                        await task_updater.add_artifact([TextPart(text=result_text)])

                await task_updater.complete()

            elif message.content:
                # If the AI just responded with text instead of calling a tool
                await task_updater.add_artifact([TextPart(text=message.content)])
                await task_updater.complete()

        except Exception as e:
            logger.error(f"🚨 LLM Processing Error: {e}")
            await task_updater.add_artifact([TextPart(text=f"Error: {str(e)}")])
            await task_updater.complete()

    def _extract_function_schema(self, func):
        """
        Dynamically extracts OpenAI-compatible JSON schema from Python functions.
        Fixes the 'string indices' error by correctly identifying 'list' as 'array'.
        """
        sig = inspect.signature(func)
        docstring = inspect.getdoc(func) or ""

        properties = {}
        required = []

        for param_name, param in sig.parameters.items():
            if param_name == 'self':
                continue

            # Default to string, but check type hints for smarter routing
            p_type = "string"
            if param.annotation == list:
                p_type = "array"
            elif param.annotation == dict:
                p_type = "object"
            elif param.annotation == int:
                p_type = "integer"
            elif param.annotation == bool:
                p_type = "boolean"

            properties[param_name] = {
                "type": p_type,
                "description": f"The {param_name} for the document generation."
            }
            required.append(param_name)

        return {
            "name": func.__name__,
            "description": docstring.split("\n")[0] if docstring else func.__name__,
            "parameters": {
                "type": "object",
                "properties": properties,
                "required": required
            },
        }

    async def execute(self, context: RequestContext, event_queue: EventQueue):
        """Standard A2A execution entry point."""
        updater = TaskUpdater(event_queue, context.task_id, context.context_id)
        await updater.submit()
        await updater.start_work()

        # Extract the user prompt from the A2A Message object
        message_text = ""
        for part in context.message.parts:
            # Handle different A2A part structures
            if hasattr(part, 'root') and hasattr(part.root, 'text'):
                message_text += part.root.text
            elif hasattr(part, 'text'):
                message_text += part.text

        await self._process_request(message_text, context, updater)

    async def cancel(self, context: RequestContext, event_queue: EventQueue):
        """Handles task cancellation requests."""
        raise ServerError(error=UnsupportedOperationError())
