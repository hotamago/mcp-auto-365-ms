"""Microsoft Teams integration module."""
from .client import TeamsClient
from .auth import TeamsAuthManager

__all__ = ["TeamsClient", "TeamsAuthManager"]
