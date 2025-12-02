"""
Exceptions for video service.
They will be eventually shift to xblocks-contrib repository once the video block is extracted.
"""


class TranscriptsGenerationException(Exception):
    pass


class NotFoundError(Exception):
    pass
