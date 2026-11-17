import gc
import os
from pathlib import Path
import sys
import tempfile
import unittest
from types import ModuleType
from unittest.mock import Mock, call, patch

import semantic_matcher
from models import ExcelRecord
from semantic_matcher import LocalSentenceEmbedder, MODEL_NAME, match_records


class LocalModelLoadingTests(unittest.TestCase):
    def setUp(self):
        self.previous_model = semantic_matcher._model_instance
        semantic_matcher._model_instance = None

    def tearDown(self):
        semantic_matcher._model_instance = self.previous_model

    def _sentence_transformers_module(self, constructor):
        module = ModuleType("sentence_transformers")
        module.SentenceTransformer = constructor
        return module

    def test_bundled_model_path_is_preferred_and_loaded_local_only(self):
        with tempfile.TemporaryDirectory() as directory:
            bundled_model = Path(directory) / "models" / "paraphrase-multilingual-MiniLM-L12-v2"
            bundled_model.mkdir(parents=True)
            model = object()
            constructor = Mock(return_value=model)
            with (
                patch.object(semantic_matcher, "BUNDLED_MODEL_DIR", bundled_model, create=True),
                patch.dict(sys.modules, {"sentence_transformers": self._sentence_transformers_module(constructor)}),
            ):
                embedder = LocalSentenceEmbedder()

        self.assertIs(model, embedder.model)
        constructor.assert_has_calls([call(str(bundled_model), device="cpu", local_files_only=True)])
        self.assertEqual(1, constructor.call_count)

    def test_cache_fallback_never_requests_network_access(self):
        with tempfile.TemporaryDirectory() as directory:
            missing_bundle = Path(directory) / "missing-model"
            constructor = Mock(side_effect=OSError("model files not found"))
            with (
                patch.object(semantic_matcher, "BUNDLED_MODEL_DIR", missing_bundle, create=True),
                patch.dict(sys.modules, {"sentence_transformers": self._sentence_transformers_module(constructor)}),
            ):
                with self.assertRaisesRegex(RuntimeError, "ExcelTracker_Offline.zip"):
                    LocalSentenceEmbedder()

        self.assertGreaterEqual(constructor.call_count, 1)
        self.assertEqual(MODEL_NAME, constructor.call_args.args[0])
        self.assertTrue(all(item.kwargs.get("local_files_only") is True for item in constructor.call_args_list))


@unittest.skipUnless(
    (Path(__file__).resolve().parents[1] / "models" / "paraphrase-multilingual-MiniLM-L12-v2").is_dir(),
    "the complete model is distributed in the offline release ZIP",
)
class BundledModelOfflineIntegrationTests(unittest.TestCase):
    def test_real_model_loads_and_semantically_matches_with_empty_huggingface_cache(self):
        import semantic_matcher

        previous_model = semantic_matcher._model_instance
        semantic_matcher._model_instance = None
        with tempfile.TemporaryDirectory() as empty_hf_cache:
            offline_environment = {
                "HF_HUB_OFFLINE": "1",
                "TRANSFORMERS_OFFLINE": "1",
                "HF_HOME": empty_hf_cache,
                "HF_HUB_CACHE": str(Path(empty_hf_cache) / "hub"),
                "TRANSFORMERS_CACHE": str(Path(empty_hf_cache) / "transformers"),
            }
            try:
                with patch.dict(os.environ, offline_environment):
                    before = [ExcelRecord(2, ("R-1", "Requirement", "系统应每100毫秒提供一次车辆速度数据。", "Open"))]
                    after = [ExcelRecord(7, ("R-1", "Requirement", "The system shall provide vehicle speed data every 100 milliseconds.", "Open"))]
                    plan = match_records(before, after, content_column=2, threshold=0.68)
            finally:
                semantic_matcher._model_instance = previous_model
                gc.collect()

        self.assertEqual(1, len(plan.matches))
        self.assertEqual("semantic", plan.matches[0].method)
        self.assertGreaterEqual(plan.matches[0].score, 0.68)

