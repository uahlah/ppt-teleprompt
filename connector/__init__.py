from .base import SlideState, BaseConnector


def get_connector(on_update):
    """Escolhe o conector correto de acordo com o sistema operacional."""
    import platform

    system = platform.system()
    if system == "Windows":
        from .windows_connector import WindowsConnector

        return WindowsConnector(on_update)
    elif system == "Darwin":
        from .mac_connector import MacConnector

        return MacConnector(on_update)
    else:
        raise RuntimeError(
            f"Sistema operacional '{system}' não suportado. "
            "Este utilitário funciona apenas com o PowerPoint desktop no Windows ou macOS."
        )


__all__ = ["SlideState", "BaseConnector", "get_connector"]
