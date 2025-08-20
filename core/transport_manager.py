"""
Advanced transport layer management for Hiel Excel MCP.
Supports multiple transport protocols: stdio, SSE, WebSocket, HTTP, and uv.
"""

import asyncio
import json
import logging
import sys
from typing import Dict, Any, Optional, Callable, List, Protocol
from dataclasses import dataclass
from enum import Enum
import websockets
import aiohttp
from aiohttp import web
import uvloop
import ssl
from contextlib import asynccontextmanager

logger = logging.getLogger(__name__)


class TransportType(Enum):
    """Supported transport types."""
    STDIO = "stdio"
    SSE = "sse"
    WEBSOCKET = "websocket"
    HTTP = "http"
    UVX = "uvx"
    GRPC = "grpc"
    PIPE = "pipe"


@dataclass
class TransportConfig:
    """Configuration for transport layer."""
    transport_type: TransportType
    host: str = "0.0.0.0"
    port: int = 8000
    ssl_cert: Optional[str] = None
    ssl_key: Optional[str] = None
    max_connections: int = 100
    connection_timeout: float = 30.0
    keepalive: bool = True
    compression: bool = True
    auth_handler: Optional[Callable] = None


class TransportInterface(Protocol):
    """Interface for transport implementations."""
    
    async def start(self) -> None:
        """Start the transport."""
        ...
    
    async def stop(self) -> None:
        """Stop the transport."""
        ...
    
    async def send(self, data: Dict[str, Any]) -> None:
        """Send data through transport."""
        ...
    
    async def receive(self) -> Dict[str, Any]:
        """Receive data from transport."""
        ...


class StdioTransport:
    """Standard I/O transport implementation."""
    
    def __init__(self, config: TransportConfig):
        self.config = config
        self._running = False
        self._reader = None
        self._writer = None
    
    async def start(self) -> None:
        """Start stdio transport."""
        self._running = True
        self._reader = asyncio.StreamReader()
        protocol = asyncio.StreamReaderProtocol(self._reader)
        
        loop = asyncio.get_event_loop()
        await loop.connect_read_pipe(lambda: protocol, sys.stdin)
        
        self._writer = sys.stdout
        logger.info("Stdio transport started")
    
    async def stop(self) -> None:
        """Stop stdio transport."""
        self._running = False
        logger.info("Stdio transport stopped")
    
    async def send(self, data: Dict[str, Any]) -> None:
        """Send data to stdout."""
        if self._writer:
            json_data = json.dumps(data) + "\n"
            self._writer.write(json_data)
            self._writer.flush()
    
    async def receive(self) -> Dict[str, Any]:
        """Receive data from stdin."""
        if self._reader:
            line = await self._reader.readline()
            if line:
                return json.loads(line.decode().strip())
        return {}


class WebSocketTransport:
    """WebSocket transport implementation."""
    
    def __init__(self, config: TransportConfig):
        self.config = config
        self._server = None
        self._clients: List[websockets.WebSocketServerProtocol] = []
        self._running = False
    
    async def start(self) -> None:
        """Start WebSocket server."""
        ssl_context = None
        if self.config.ssl_cert and self.config.ssl_key:
            ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            ssl_context.load_cert_chain(self.config.ssl_cert, self.config.ssl_key)
        
        self._server = await websockets.serve(
            self._handle_client,
            self.config.host,
            self.config.port,
            ssl=ssl_context,
            compression="deflate" if self.config.compression else None,
            max_size=10**7,  # 10MB max message size
            ping_interval=20 if self.config.keepalive else None,
            ping_timeout=10 if self.config.keepalive else None
        )
        self._running = True
        logger.info(f"WebSocket transport started on {self.config.host}:{self.config.port}")
    
    async def stop(self) -> None:
        """Stop WebSocket server."""
        self._running = False
        if self._server:
            self._server.close()
            await self._server.wait_closed()
        
        # Close all client connections
        for client in self._clients:
            await client.close()
        self._clients.clear()
        
        logger.info("WebSocket transport stopped")
    
    async def _handle_client(self, websocket, path):
        """Handle WebSocket client connection."""
        # Authentication if configured
        if self.config.auth_handler:
            try:
                auth_result = await self.config.auth_handler(websocket)
                if not auth_result:
                    await websocket.close(1008, "Authentication failed")
                    return
            except Exception as e:
                logger.error(f"Authentication error: {e}")
                await websocket.close(1011, "Authentication error")
                return
        
        self._clients.append(websocket)
        logger.info(f"WebSocket client connected from {websocket.remote_address}")
        
        try:
            async for message in websocket:
                # Process incoming message
                data = json.loads(message)
                await self._process_message(data, websocket)
        except websockets.exceptions.ConnectionClosed:
            logger.info(f"WebSocket client disconnected: {websocket.remote_address}")
        except Exception as e:
            logger.error(f"WebSocket error: {e}")
        finally:
            if websocket in self._clients:
                self._clients.remove(websocket)
    
    async def _process_message(self, data: Dict[str, Any], websocket) -> None:
        """Process incoming WebSocket message."""
        # Override in subclass to handle messages
        pass
    
    async def send(self, data: Dict[str, Any], client=None) -> None:
        """Send data to WebSocket clients."""
        message = json.dumps(data)
        
        if client:
            # Send to specific client
            await client.send(message)
        else:
            # Broadcast to all clients
            disconnected = []
            for client in self._clients:
                try:
                    await client.send(message)
                except websockets.exceptions.ConnectionClosed:
                    disconnected.append(client)
            
            # Remove disconnected clients
            for client in disconnected:
                if client in self._clients:
                    self._clients.remove(client)
    
    async def receive(self) -> Dict[str, Any]:
        """Not implemented for WebSocket (uses callback pattern)."""
        raise NotImplementedError("WebSocket uses callback pattern")


class SSETransport:
    """Server-Sent Events transport implementation."""
    
    def __init__(self, config: TransportConfig):
        self.config = config
        self._app = None
        self._runner = None
        self._site = None
        self._clients: List[web.StreamResponse] = []
    
    async def start(self) -> None:
        """Start SSE server."""
        self._app = web.Application()
        self._app.router.add_get('/events', self._handle_sse)
        self._app.router.add_post('/send', self._handle_send)
        
        self._runner = web.AppRunner(self._app)
        await self._runner.setup()
        
        ssl_context = None
        if self.config.ssl_cert and self.config.ssl_key:
            ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            ssl_context.load_cert_chain(self.config.ssl_cert, self.config.ssl_key)
        
        self._site = web.TCPSite(
            self._runner,
            self.config.host,
            self.config.port,
            ssl_context=ssl_context
        )
        await self._site.start()
        
        logger.info(f"SSE transport started on {self.config.host}:{self.config.port}")
    
    async def stop(self) -> None:
        """Stop SSE server."""
        # Close all client connections
        for client in self._clients:
            await client.write_eof()
        self._clients.clear()
        
        if self._runner:
            await self._runner.cleanup()
        
        logger.info("SSE transport stopped")
    
    async def _handle_sse(self, request):
        """Handle SSE client connection."""
        response = web.StreamResponse(
            status=200,
            headers={
                'Content-Type': 'text/event-stream',
                'Cache-Control': 'no-cache',
                'Connection': 'keep-alive',
                'Access-Control-Allow-Origin': '*'
            }
        )
        await response.prepare(request)
        
        self._clients.append(response)
        logger.info(f"SSE client connected from {request.remote}")
        
        try:
            # Keep connection alive
            while True:
                await asyncio.sleep(30)
                await response.write(b': ping\n\n')  # SSE comment for keepalive
        except Exception as e:
            logger.error(f"SSE error: {e}")
        finally:
            if response in self._clients:
                self._clients.remove(response)
        
        return response
    
    async def _handle_send(self, request):
        """Handle incoming data via POST."""
        data = await request.json()
        # Process the data (override in subclass)
        return web.json_response({'status': 'received'})
    
    async def send(self, data: Dict[str, Any]) -> None:
        """Send data to SSE clients."""
        event_data = f"data: {json.dumps(data)}\n\n"
        
        disconnected = []
        for client in self._clients:
            try:
                await client.write(event_data.encode())
            except Exception:
                disconnected.append(client)
        
        # Remove disconnected clients
        for client in disconnected:
            if client in self._clients:
                self._clients.remove(client)
    
    async def receive(self) -> Dict[str, Any]:
        """Not implemented for SSE (uses callback pattern)."""
        raise NotImplementedError("SSE uses callback pattern")


class UvTransport:
    """Uvloop-optimized transport implementation."""
    
    def __init__(self, config: TransportConfig):
        self.config = config
        self._server = None
        self._loop = None
    
    async def start(self) -> None:
        """Start UV transport with optimized event loop."""
        # Install uvloop for better performance
        uvloop.install()
        self._loop = asyncio.get_event_loop()
        
        # Create optimized server
        self._server = await self._loop.create_server(
            lambda: UvProtocol(self),
            self.config.host,
            self.config.port,
            reuse_address=True,
            reuse_port=True  # SO_REUSEPORT for load balancing
        )
        
        logger.info(f"UV transport started on {self.config.host}:{self.config.port}")
    
    async def stop(self) -> None:
        """Stop UV transport."""
        if self._server:
            self._server.close()
            await self._server.wait_closed()
        logger.info("UV transport stopped")
    
    async def send(self, data: Dict[str, Any]) -> None:
        """Send data through UV transport."""
        # Implementation depends on protocol
        pass
    
    async def receive(self) -> Dict[str, Any]:
        """Receive data from UV transport."""
        # Implementation depends on protocol
        pass


class UvProtocol(asyncio.Protocol):
    """Protocol implementation for UV transport."""
    
    def __init__(self, transport_manager):
        self.transport_manager = transport_manager
        self.transport = None
        self.buffer = b''
    
    def connection_made(self, transport):
        """Handle new connection."""
        self.transport = transport
        peername = transport.get_extra_info('peername')
        logger.info(f"UV connection from {peername}")
    
    def data_received(self, data):
        """Handle received data."""
        self.buffer += data
        # Process complete messages
        while b'\n' in self.buffer:
            line, self.buffer = self.buffer.split(b'\n', 1)
            try:
                message = json.loads(line.decode())
                # Process message asynchronously
                asyncio.create_task(self._process_message(message))
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON received: {e}")
    
    async def _process_message(self, message: Dict[str, Any]):
        """Process received message."""
        # Override in subclass
        pass
    
    def connection_lost(self, exc):
        """Handle connection loss."""
        if exc:
            logger.error(f"UV connection lost: {exc}")
        else:
            logger.info("UV connection closed")


class TransportManager:
    """Manages multiple transport layers."""
    
    def __init__(self):
        self._transports: Dict[TransportType, TransportInterface] = {}
        self._active_transports: List[TransportType] = []
        self._default_transport: Optional[TransportType] = None
        self._message_handlers: Dict[str, Callable] = {}
    
    def register_transport(self, transport_type: TransportType, 
                         transport: TransportInterface) -> None:
        """Register a transport implementation."""
        self._transports[transport_type] = transport
        logger.info(f"Registered transport: {transport_type.value}")
    
    def register_handler(self, message_type: str, handler: Callable) -> None:
        """Register a message handler."""
        self._message_handlers[message_type] = handler
    
    async def start_transport(self, transport_type: TransportType) -> None:
        """Start a specific transport."""
        if transport_type not in self._transports:
            raise ValueError(f"Transport {transport_type.value} not registered")
        
        transport = self._transports[transport_type]
        await transport.start()
        self._active_transports.append(transport_type)
        
        if not self._default_transport:
            self._default_transport = transport_type
    
    async def stop_transport(self, transport_type: TransportType) -> None:
        """Stop a specific transport."""
        if transport_type in self._active_transports:
            transport = self._transports[transport_type]
            await transport.stop()
            self._active_transports.remove(transport_type)
            
            if self._default_transport == transport_type:
                self._default_transport = self._active_transports[0] if self._active_transports else None
    
    async def stop_all(self) -> None:
        """Stop all active transports."""
        for transport_type in list(self._active_transports):
            await self.stop_transport(transport_type)
    
    async def send(self, data: Dict[str, Any], 
                  transport_type: Optional[TransportType] = None) -> None:
        """Send data through transport."""
        target = transport_type or self._default_transport
        if not target or target not in self._active_transports:
            raise RuntimeError("No active transport available")
        
        transport = self._transports[target]
        await transport.send(data)
    
    async def broadcast(self, data: Dict[str, Any]) -> None:
        """Broadcast data to all active transports."""
        tasks = []
        for transport_type in self._active_transports:
            transport = self._transports[transport_type]
            tasks.append(transport.send(data))
        
        if tasks:
            await asyncio.gather(*tasks, return_exceptions=True)
    
    @asynccontextmanager
    async def transport_context(self, transport_type: TransportType):
        """Context manager for transport lifecycle."""
        try:
            await self.start_transport(transport_type)
            yield self._transports[transport_type]
        finally:
            await self.stop_transport(transport_type)
    
    def get_stats(self) -> Dict[str, Any]:
        """Get transport statistics."""
        return {
            'registered': list(self._transports.keys()),
            'active': self._active_transports,
            'default': self._default_transport,
            'handlers': list(self._message_handlers.keys())
        }


# Global transport manager instance
transport_manager = TransportManager()
