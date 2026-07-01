from typing import TypedDict, Literal, Optional

class ParagraphLocator(TypedDict):
    type: Literal["paragraph"]
    index: int

class HeadingLocator(TypedDict):
    type: Literal["heading"]
    text: str

class BookmarkLocator(TypedDict):
    type: Literal["bookmark"]
    name: str

class SearchLocator(TypedDict):
    type: Literal["search"]
    text: str
    instance: Optional[int]

class Margins(TypedDict, total=False):
    top: float
    bottom: float
    left: float
    right: float
