# Workbook Context Management and Caching System

The workbook context management system provides intelligent caching and lifecycle management for Excel workbooks to optimize performance when performing multiple operations on the same files.

## Key Features

- **Intelligent Caching**: Automatically caches workbooks to avoid repeated loading
- **Thread Safety**: Safe for concurrent access from multiple threads
- **Automatic Lifecycle Management**: Handles workbook opening, saving, and closing
- **Memory Management**: LRU eviction and expiration-based cleanup
- **Performance Optimization**: Significant performance improvements for repeated operations
- **Context Manager Support**: Clean, Pythonic API with automatic resource management

## Core Components

### WorkbookContext

The `WorkbookContext` class manages individual workbook instances with caching and lifecycle management.

```python
from hiel_excel_mcp.core.workbook_context import WorkbookContext

# Create context for a workbook
context = WorkbookContext('path/to/file.xlsx', read_only=False, data_only=False)

# Use as context manager
with context as workbook:
    worksheet = workbook.active
    worksheet['A1'] = 'Hello World'
# Workbook is automatically saved on exit (if not read-only)
```

**Key Features:**
- Automatic workbook loading and creation
- Thread-safe operations with RLock
- Dirty tracking for efficient saving
- Access count and timestamp tracking
- Expiration detection

### WorkbookCache

The `WorkbookCache` class manages multiple workbook contexts with intelligent caching strategies.

```python
from hiel_excel_mcp.core.workbook_context import WorkbookCache

# Create cache with custom settings
cache = WorkbookCache(max_size=10, max_age_seconds=300)

# Get cached context
context = cache.get_context('file.xlsx', read_only=False)
```

**Key Features:**
- LRU (Least Recently Used) eviction
- Time-based expiration
- Thread-safe operations
- Cache statistics and monitoring
- Automatic cleanup of expired contexts

### Global Convenience Functions

The module provides convenient global functions for common operations:

```python
from hiel_excel_mcp.core.workbook_context import (
    workbook_context, get_cache_stats, invalidate_cache, 
    clear_cache, configure_cache
)

# Use workbook with automatic caching
with workbook_context('file.xlsx') as wb:
    ws = wb.active
    ws['A1'] = 'Data'

# Get cache statistics
stats = get_cache_stats()
print(f"Cache hit rate: {stats['hit_rate']:.2%}")

# Invalidate specific file
invalidate_cache('file.xlsx')

# Clear entire cache
clear_cache()

# Configure cache settings
configure_cache(max_size=20, max_age_seconds=600)
```

## Performance Benefits

### Before (Without Caching)
```python
# Each operation loads the workbook from disk
for i in range(10):
    wb = load_workbook('large_file.xlsx')
    ws = wb.active
    ws[f'A{i}'] = f'Value {i}'
    wb.save('large_file.xlsx')
    wb.close()
# Result: 10 disk reads, 10 disk writes
```

### After (With Workbook Context)
```python
# Workbook is loaded once and reused
for i in range(10):
    with workbook_context('large_file.xlsx') as wb:
        ws = wb.active
        ws[f'A{i}'] = f'Value {i}'
# Result: 1 disk read, 1 disk write (at the end)
```

## Cache Configuration

### Default Settings
- **Max Size**: 10 workbooks
- **Max Age**: 300 seconds (5 minutes)
- **Thread Safe**: Yes
- **Auto Cleanup**: Yes

### Customization
```python
# Configure for high-performance scenarios
configure_cache(
    max_size=50,        # Cache up to 50 workbooks
    max_age_seconds=1800  # Keep for 30 minutes
)

# Configure for memory-constrained environments
configure_cache(
    max_size=3,         # Cache only 3 workbooks
    max_age_seconds=60  # Keep for 1 minute
)
```

## Thread Safety

The workbook context system is fully thread-safe:

```python
import threading
from hiel_excel_mcp.core.workbook_context import workbook_context

def worker(worker_id):
    with workbook_context('shared_file.xlsx') as wb:
        ws = wb.active
        ws[f'A{worker_id}'] = f'Worker {worker_id}'

# Multiple threads can safely access the same file
threads = [threading.Thread(target=worker, args=(i,)) for i in range(5)]
for thread in threads:
    thread.start()
for thread in threads:
    thread.join()
```

## Integration with Base Tools

The workbook context system integrates seamlessly with the base tool infrastructure:

```python
from hiel_excel_mcp.core.base_tool import BaseTool, operation_route
from hiel_excel_mcp.core.workbook_context import workbook_context

class ExampleTool(BaseTool):
    def get_tool_name(self):
        return 'example_tool'
    
    def get_tool_description(self):
        return 'Example tool using workbook context'
    
    @operation_route('write_data', 'Write data to workbook', ['filepath', 'data'])
    def write_data(self, filepath, data, **kwargs):
        with workbook_context(filepath) as wb:
            ws = wb.active
            ws['A1'] = data
        return {'success': True, 'data_written': data}
```

## Cache Statistics and Monitoring

Monitor cache performance with detailed statistics:

```python
stats = get_cache_stats()
print(f"""
Cache Statistics:
- Size: {stats['size']}/{stats['max_size']} workbooks
- Hit Rate: {stats['hit_rate']:.2%}
- Total Accesses: {stats['total_accesses']}
- Hits: {stats['hits']}
- Misses: {stats['misses']}
- Evictions: {stats['evictions']}
""")

# Detailed context information
for context_info in stats['contexts']:
    print(f"File: {context_info['filepath']}")
    print(f"  Access Count: {context_info['access_count']}")
    print(f"  Is Dirty: {context_info['is_dirty']}")
    print(f"  Is Loaded: {context_info['is_loaded']}")
```

## Error Handling

The system provides robust error handling:

```python
try:
    with workbook_context('nonexistent.xlsx') as wb:
        # If file doesn't exist, a new workbook is created
        ws = wb.active
        ws['A1'] = 'New file'
except Exception as e:
    print(f"Error: {e}")

# Read-only mode prevents accidental modifications
try:
    context = WorkbookContext('file.xlsx', read_only=True)
    context.save()  # Raises ValueError
except ValueError as e:
    print(f"Cannot save read-only workbook: {e}")
```

## Best Practices

### 1. Use Context Managers
Always use context managers for automatic resource management:
```python
# Good
with workbook_context('file.xlsx') as wb:
    # Work with workbook
    pass

# Avoid
context = WorkbookContext('file.xlsx')
wb = context.get_workbook()
# Manual cleanup required
```

### 2. Configure Cache Appropriately
Adjust cache settings based on your use case:
```python
# For batch processing with many files
configure_cache(max_size=100, max_age_seconds=3600)

# For interactive applications
configure_cache(max_size=5, max_age_seconds=300)
```

### 3. Monitor Cache Performance
Regularly check cache statistics to optimize settings:
```python
stats = get_cache_stats()
if stats['hit_rate'] < 0.5:  # Less than 50% hit rate
    # Consider increasing cache size or age
    configure_cache(max_size=stats['max_size'] * 2)
```

### 4. Handle Large Files Carefully
For very large files, consider using read-only mode when possible:
```python
# Read-only for analysis
with workbook_context('large_file.xlsx', read_only=True) as wb:
    # Analyze data without loading formulas
    pass

# Data-only mode for faster loading
with workbook_context('large_file.xlsx', data_only=True) as wb:
    # Load only values, not formulas
    pass
```

## Requirements Satisfied

This implementation satisfies the following requirements:

- **6.1**: Performance optimization through workbook reuse and intelligent caching
- **6.2**: Efficient memory management with LRU eviction and expiration
- **6.3**: Maintains or improves performance compared to original server through caching

The workbook context management system provides a solid foundation for the optimized Excel MCP server, enabling efficient handling of multiple operations on the same files while maintaining thread safety and proper resource management.