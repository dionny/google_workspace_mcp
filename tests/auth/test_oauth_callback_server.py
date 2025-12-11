"""Tests for OAuth callback server auto port detection."""

from unittest.mock import patch, MagicMock

from auth.oauth_callback_server import MinimalOAuthServer


class TestMinimalOAuthServerPortDetection:
    """Test automatic port detection in MinimalOAuthServer."""

    def test_find_available_port_preferred_available(self):
        """When preferred port is available, should return it."""
        server = MinimalOAuthServer(port=8000)

        # Mock socket to indicate port is available
        with patch("socket.socket") as mock_socket:
            mock_instance = MagicMock()
            mock_socket.return_value.__enter__.return_value = mock_instance
            mock_instance.bind.return_value = None  # No error = port available

            result = server._find_available_port("localhost", 8000)

        assert result == 8000

    def test_find_available_port_preferred_in_use(self):
        """When preferred port is in use, should find another port."""
        server = MinimalOAuthServer(port=8000)

        # Track which ports have been tried
        tried_ports = []

        def mock_bind(addr):
            host, port = addr
            tried_ports.append(port)
            if port == 8000:
                raise OSError("Port in use")
            # Port 8001 is available

        with patch("socket.socket") as mock_socket:
            mock_instance = MagicMock()
            mock_socket.return_value.__enter__.return_value = mock_instance
            mock_instance.bind.side_effect = mock_bind

            result = server._find_available_port("localhost", 8000)

        assert result == 8001
        assert 8000 in tried_ports

    def test_find_available_port_range_exhausted(self):
        """When all ports in range are in use, should return None."""
        server = MinimalOAuthServer(port=8000)

        # All ports are in use
        def mock_bind(addr):
            raise OSError("Port in use")

        with patch("socket.socket") as mock_socket:
            mock_instance = MagicMock()
            mock_socket.return_value.__enter__.return_value = mock_instance
            mock_instance.bind.side_effect = mock_bind

            result = server._find_available_port("localhost", 8000)

        assert result is None

    def test_port_range_constants(self):
        """Verify port range constants are reasonable."""
        assert MinimalOAuthServer.PORT_RANGE_START == 8000
        assert MinimalOAuthServer.PORT_RANGE_END == 8100
        assert MinimalOAuthServer.PORT_RANGE_END > MinimalOAuthServer.PORT_RANGE_START

    def test_init_stores_configured_port(self):
        """Should store configured port separately from actual port."""
        server = MinimalOAuthServer(port=8042)

        assert server.configured_port == 8042
        assert server.port == 8042  # Initially same


class TestMinimalOAuthServerStart:
    """Test the start method with port detection."""

    def test_start_returns_tuple_with_port(self):
        """Start should return a 3-tuple including the actual port."""
        server = MinimalOAuthServer(port=8000)

        with patch.object(
            server, "_find_available_port", return_value=8005
        ), patch.object(server, "_update_oauth_config_port"), patch(
            "threading.Thread"
        ) as mock_thread, patch(
            "socket.socket"
        ) as mock_socket:
            # Mock thread start
            mock_thread_instance = MagicMock()
            mock_thread.return_value = mock_thread_instance

            # Mock socket connect to simulate server started
            mock_socket_instance = MagicMock()
            mock_socket.return_value.__enter__.return_value = mock_socket_instance
            mock_socket_instance.connect_ex.return_value = 0

            success, error, port = server.start()

        assert success is True
        assert error == ""
        assert port == 8005

    def test_start_no_available_port(self):
        """When no port is available, should return error."""
        server = MinimalOAuthServer(port=8000)

        with patch.object(server, "_find_available_port", return_value=None):
            success, error, port = server.start()

        assert success is False
        assert "No available ports" in error
        assert port is None

    def test_start_updates_oauth_config_when_port_changes(self):
        """When port changes, should update OAuth config."""
        server = MinimalOAuthServer(port=8000)

        with patch.object(
            server, "_find_available_port", return_value=8005
        ), patch.object(
            server, "_update_oauth_config_port"
        ) as mock_update, patch(
            "threading.Thread"
        ) as mock_thread, patch(
            "socket.socket"
        ) as mock_socket:
            mock_thread_instance = MagicMock()
            mock_thread.return_value = mock_thread_instance
            mock_socket_instance = MagicMock()
            mock_socket.return_value.__enter__.return_value = mock_socket_instance
            mock_socket_instance.connect_ex.return_value = 0

            server.start()

        mock_update.assert_called_once_with(8005)

    def test_start_does_not_update_config_when_port_unchanged(self):
        """When port doesn't change, should not update OAuth config."""
        server = MinimalOAuthServer(port=8000)

        with patch.object(
            server, "_find_available_port", return_value=8000
        ), patch.object(
            server, "_update_oauth_config_port"
        ) as mock_update, patch(
            "threading.Thread"
        ) as mock_thread, patch(
            "socket.socket"
        ) as mock_socket:
            mock_thread_instance = MagicMock()
            mock_thread.return_value = mock_thread_instance
            mock_socket_instance = MagicMock()
            mock_socket.return_value.__enter__.return_value = mock_socket_instance
            mock_socket_instance.connect_ex.return_value = 0

            server.start()

        mock_update.assert_not_called()
