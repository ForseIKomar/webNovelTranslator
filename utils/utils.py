from utils.parser.allnovelbook import *
from utils.parser.fucknovelpia import *

parsers = {
    'allnovelbook.com': allnovelbook_extract_title,
    'fucknovelpia.com': fucknovelpia_extract_title
}
