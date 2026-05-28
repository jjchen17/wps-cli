"""WPS COM 后端测试"""


class TestMockBackend:
    """测试模拟后端"""

    def test_connect(self, mock_backend):
        app = mock_backend.connect("writer")
        assert app.Name == "Mock writer"

    def test_is_alive(self, mock_backend, mock_app):
        assert mock_backend.is_alive(mock_app) is True

    def test_get_version(self, mock_backend, mock_app):
        assert mock_backend.get_version(mock_app) == "12.0.0-test"

    def test_disconnect(self, mock_backend, mock_app):
        mock_backend.disconnect(mock_app)  # 不应抛异常
