"""Load YAML with PyYAML, rejecting duplicate mapping keys.

PyYAML keeps the last value of a duplicate key without an error, while GitHub
may read a different one, so a check could pass on a value GitHub does not
use. The workflow control tests and Repository integrity read the workflows,
Dependabot's configuration and the issue forms through this loader.
"""

from typing import Any

import yaml


# PyYAML has no type information, so mypy sees its loader as Any.
class StrictLoader(yaml.SafeLoader):  # type: ignore[misc]
    """A safe loader whose mappings fail on a repeated key."""


def construct_unique_mapping(loader: StrictLoader, node: Any,
                             deep: bool = False) -> Any:
    loader.flatten_mapping(node)
    seen: set[Any] = set()
    for key_node, _ in node.value:
        key = loader.construct_object(key_node, deep=deep)
        if key in seen:
            raise yaml.constructor.ConstructorError(
                'while constructing a mapping', node.start_mark,
                f'found duplicate key {key!r}', key_node.start_mark)
        seen.add(key)
    return loader.construct_mapping(node, deep=deep)


StrictLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
                             construct_unique_mapping)


def load(text: str) -> Any:
    """Parse one YAML document; a duplicate key raises yaml.YAMLError."""
    return yaml.load(text, Loader=StrictLoader)
