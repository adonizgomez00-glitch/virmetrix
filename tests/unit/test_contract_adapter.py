"""Contract test: both adapters implement the EvaluationAdapter port."""

from virmetrix.domain.ports import EvaluationAdapter
from virmetrix.infrastructure.evaluators.portable import PortableDocxEvaluator
from virmetrix.infrastructure.evaluators.word_com import WordComAdapter


def test_portable_is_adapter():
    assert isinstance(PortableDocxEvaluator(), EvaluationAdapter)


def test_wordcom_is_adapter():
    assert isinstance(WordComAdapter(), EvaluationAdapter)
