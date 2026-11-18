from __future__ import annotations


class VisibleAnswerError(ValueError):
    pass


class LeadingThoughtFilter:
    def __init__(self, *, max_hidden_characters=32000):
        if type(max_hidden_characters) is not int or max_hidden_characters < 1:
            raise ValueError('thought_filter_limit_invalid')
        self._limit = max_hidden_characters
        self._spent = 0
        self._buffer = ''
        self._mode = 'leading'
        self._finished = False
        self._failed = False

    def _reject(self, reason):
        self._failed = True
        self._buffer = ''
        raise VisibleAnswerError(reason)

    def _budget(self, amount, pending=0):
        self._spent += amount
        if self._spent + pending > self._limit:
            self._reject('leading_thought_limit')

    def feed(self, text):
        if self._failed or self._finished:
            raise VisibleAnswerError('thought_filter_closed')
        if type(text) is not str:
            self._reject('answer_text_invalid')
        if self._mode == 'visible':
            return text
        data, self._buffer = self._buffer + text, ''
        opening, closing = '<think>', '</think>'
        while True:
            if self._mode == 'leading':
                stripped = data.lstrip()
                if not stripped or (opening.startswith(stripped) and stripped != opening):
                    self._budget(0, len(data))
                    self._buffer = data
                    return ''
                if stripped.startswith(opening):
                    self._budget(len(data) - len(stripped) + len(opening))
                    data = stripped[len(opening):]
                    self._mode = 'thought'
                else:
                    self._mode = 'visible'
                    return data
            end = data.find(closing)
            nested = data.find(opening)
            if nested >= 0 and (end < 0 or nested < end):
                self._reject('nested_leading_thought')
            if end < 0:
                retained = min(len(data), len(closing) - 1)
                self._budget(len(data) - retained, retained)
                self._buffer = data[-retained:] if retained else ''
                return ''
            self._budget(end + len(closing))
            data = data[end + len(closing):]
            self._mode = 'leading'

    def finish(self):
        if self._failed or self._finished:
            raise VisibleAnswerError('thought_filter_closed')
        if self._mode == 'thought':
            self._reject('unclosed_leading_thought')
        if self._buffer.lstrip():
            self._reject('incomplete_leading_thought_prefix')
        self._finished = True
        tail, self._buffer = self._buffer, ''
        return tail


def visible_answer(text):
    stream = LeadingThoughtFilter()
    return stream.feed(text) + stream.finish()
