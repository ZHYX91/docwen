"""Application ports (interfaces): runtime port, config port, presenter port.

These ports define WHAT the application layer needs from the outside world.
The runtime/bundle layers provide the IMPLEMENTATION of these ports.

Dependency Inversion Principle:
    application defines the interface → runtime/bundle implements it.
    NOT: application imports runtime to get its interface.
"""

from docwen_application.ports.runtime import (
    ArtifactBundleCommitPort,
    CapabilityDiscoveryPort,
    ConfigPort,
    OutputManifestPersistencePort,
    PresenterPort,
    RuntimePort,
)

__all__ = [
    "ArtifactBundleCommitPort",
    "CapabilityDiscoveryPort",
    "ConfigPort",
    "OutputManifestPersistencePort",
    "PresenterPort",
    "RuntimePort",
]
