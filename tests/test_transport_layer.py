"""
Comprehensive tests for transport layer implementations.
"""

import pytest
import asyncio
import json
import websockets
from unittest.mock import Mock, patch, AsyncMock
from datetime import datetime

from core.transport_manager import (
    TransportManager, TransportType, TransportConfig,
    StdioTransport, WebSocketTransport, SSETransport, 
    UvTransport
)


class TestTransportManager:
    """Test TransportManager functionality."""
    
    @pytest.fixture
    def manager(self):
        """Create a fresh TransportManager instance."""
        return TransportManager()
    
    @pytest.fixture
    def config(self):
        """Create default transport config."""
        return TransportConfig(
            transport_type=TransportType.STDIO,
            host="localhost",
            port=8000
        )
    
    @pytest.mark.asyncio
    async def test_register_transport(self, manager):
        """Test transport registration."""
        config = TransportConfig(transport_type=TransportType.STDIO)
        transport = StdioTransport(config)
        
        manager.register_transport(TransportType.STDIO, transport)
        assert TransportType.STDIO in manager._transports
    
    @pytest.mark.asyncio
    async def test_start_stop_transport(self, manager, config):
        """Test starting and stopping transports."""
        transport = AsyncMock()
        manager.register_transport(TransportType.STDIO, transport)
        
        await manager.start_transport(TransportType.STDIO)
        transport.start.assert_called_once()
        assert TransportType.STDIO in manager._active_transports
        
        await manager.stop_transport(TransportType.STDIO)
        transport.stop.assert_called_once()
        assert TransportType.STDIO not in manager._active_transports
    
    @pytest.mark.asyncio
    async def test_send_data(self, manager):
        """Test sending data through transport."""
        transport = AsyncMock()
        manager.register_transport(TransportType.STDIO, transport)
        await manager.start_transport(TransportType.STDIO)
        
        test_data = {"message": "test", "timestamp": datetime.now().isoformat()}
        await manager.send(test_data)
        
        transport.send.assert_called_once_with(test_data)
    
    @pytest.mark.asyncio
    async def test_broadcast_data(self, manager):
        """Test broadcasting to multiple transports."""
        stdio_transport = AsyncMock()
        ws_transport = AsyncMock()
        
        manager.register_transport(TransportType.STDIO, stdio_transport)
        manager.register_transport(TransportType.WEBSOCKET, ws_transport)
        
        await manager.start_transport(TransportType.STDIO)
        await manager.start_transport(TransportType.WEBSOCKET)
        
        test_data = {"broadcast": "test"}
        await manager.broadcast(test_data)
        
        stdio_transport.send.assert_called_once_with(test_data)
        ws_transport.send.assert_called_once_with(test_data)
    
    @pytest.mark.asyncio
    async def test_transport_context_manager(self, manager, config):
        """Test transport context manager."""
        transport = AsyncMock()
        manager.register_transport(TransportType.STDIO, transport)
        
        async with manager.transport_context(TransportType.STDIO) as t:
            assert t == transport
            transport.start.assert_called_once()
        
        transport.stop.assert_called_once()
    
    def test_get_stats(self, manager):
        """Test getting transport statistics."""
        manager.register_transport(TransportType.STDIO, Mock())
        manager._active_transports = [TransportType.STDIO]
        manager._default_transport = TransportType.STDIO
        
        stats = manager.get_stats()
        assert stats['registered'] == [TransportType.STDIO]
        assert stats['active'] == [TransportType.STDIO]
        assert stats['default'] == TransportType.STDIO


class TestStdioTransport:
    """Test StdioTransport implementation."""
    
    @pytest.fixture
    def transport(self):
        """Create StdioTransport instance."""
        config = TransportConfig(transport_type=TransportType.STDIO)
        return StdioTransport(config)
    
    @pytest.mark.asyncio
    @patch('asyncio.get_event_loop')
    async def test_start(self, mock_loop, transport):
        """Test starting stdio transport."""
        mock_loop.return_value.connect_read_pipe = AsyncMock()
        
        await transport.start()
        assert transport._running is True
        assert transport._reader is not None
    
    @pytest.mark.asyncio
    async def test_stop(self, transport):
        """Test stopping stdio transport."""
        transport._running = True
        await transport.stop()
        assert transport._running is False
    
    @pytest.mark.asyncio
    @patch('sys.stdout')
    async def test_send(self, mock_stdout, transport):
        """Test sending data to stdout."""
        transport._writer = mock_stdout
        test_data = {"test": "data"}
        
        await transport.send(test_data)
        
        expected_output = json.dumps(test_data) + "\n"
        mock_stdout.write.assert_called_once_with(expected_output)
        mock_stdout.flush.assert_called_once()


class TestWebSocketTransport:
    """Test WebSocketTransport implementation."""
    
    @pytest.fixture
    def transport(self):
        """Create WebSocketTransport instance."""
        config = TransportConfig(
            transport_type=TransportType.WEBSOCKET,
            host="localhost",
            port=8001
        )
        return WebSocketTransport(config)
    
    @pytest.mark.asyncio
    @patch('websockets.serve')
    async def test_start(self, mock_serve, transport):
        """Test starting WebSocket server."""
        # Create a proper async context manager mock
        mock_server = AsyncMock()
        mock_server.__aenter__ = AsyncMock(return_value=mock_server)
        mock_server.__aexit__ = AsyncMock(return_value=None)
        
        # Make serve return a coroutine that returns the mock server
        async def mock_serve_coro(*args, **kwargs):
            return mock_server
        
        mock_serve.side_effect = mock_serve_coro
        
        await transport.start()
        
        assert transport._running is True
        assert transport._server == mock_server
        mock_serve.assert_called_once()
    
    @pytest.mark.asyncio
    async def test_stop(self, transport):
        """Test stopping WebSocket server."""
        mock_server = AsyncMock()
        transport._server = mock_server
        transport._running = True
        
        mock_client = AsyncMock()
        transport._clients = [mock_client]
        
        await transport.stop()
        
        assert transport._running is False
        mock_server.close.assert_called_once()
        mock_client.close.assert_called_once()
        assert len(transport._clients) == 0
    
    @pytest.mark.asyncio
    async def test_send_to_all_clients(self, transport):
        """Test broadcasting to all WebSocket clients."""
        client1 = AsyncMock()
        client2 = AsyncMock()
        transport._clients = [client1, client2]
        
        test_data = {"message": "broadcast"}
        await transport.send(test_data)
        
        expected_message = json.dumps(test_data)
        client1.send.assert_called_once_with(expected_message)
        client2.send.assert_called_once_with(expected_message)
    
    @pytest.mark.asyncio
    async def test_send_to_specific_client(self, transport):
        """Test sending to specific WebSocket client."""
        client1 = AsyncMock()
        client2 = AsyncMock()
        transport._clients = [client1, client2]
        
        test_data = {"message": "specific"}
        await transport.send(test_data, client=client1)
        
        expected_message = json.dumps(test_data)
        client1.send.assert_called_once_with(expected_message)
        client2.send.assert_not_called()
    
    @pytest.mark.asyncio
    async def test_handle_disconnected_client(self, transport):
        """Test handling disconnected clients during send."""
        client1 = AsyncMock()
        client2 = AsyncMock()
        client1.send.side_effect = websockets.exceptions.ConnectionClosed(None, None)
        transport._clients = [client1, client2]
        
        test_data = {"message": "test"}
        await transport.send(test_data)
        
        # Client1 should be removed, client2 should receive message
        assert client1 not in transport._clients
        assert client2 in transport._clients
        client2.send.assert_called_once()
