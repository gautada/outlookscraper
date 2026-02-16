#!/usr/bin/env python3
"""
Standalone mTLS POST test script.
Tests client certificate authentication with verbose logging.

Usage:
    uv run python test_mtls.py <url> --ca <ca.pem> --cert <crt.pem> --key <key.pem>
    
Example:
    uv run python test_mtls.py https://api.example.com/endpoint \
        --ca ~/.config/cauth/ca.pem \
        --cert ~/.config/cauth/crt.pem \
        --key ~/.config/cauth/key.pem
"""

import argparse
import json
import logging
import ssl
import sys
from datetime import datetime, timezone
from pathlib import Path

import httpx

# Configure verbose logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s | %(levelname)-8s | %(name)s | %(message)s',
    datefmt='%H:%M:%S'
)

# Enable httpx and httpcore debug logging
logging.getLogger("httpx").setLevel(logging.DEBUG)
logging.getLogger("httpcore").setLevel(logging.DEBUG)
logging.getLogger("ssl").setLevel(logging.DEBUG)

logger = logging.getLogger("mtls_test")


def create_ssl_context(
    ca_path: Path,
    cert_path: Path,
    key_path: Path,
    include_system_cas: bool = True
) -> ssl.SSLContext:
    """
    Create SSL context for mTLS.
    
    Args:
        ca_path: Custom CA certificate file
        cert_path: Client certificate file
        key_path: Client private key file
        include_system_cas: If True, also trust system CA bundle (for public servers)
    """
    logger.info("Creating SSL context...")
    logger.debug(f"  CA file:   {ca_path}")
    logger.debug(f"  Cert file: {cert_path}")
    logger.debug(f"  Key file:  {key_path}")
    logger.debug(f"  Include system CAs: {include_system_cas}")
    
    # Verify files exist
    for path, name in [(ca_path, "CA"), (cert_path, "cert"), (key_path, "key")]:
        if not path.exists():
            logger.error(f"{name} file not found: {path}")
            raise FileNotFoundError(f"{name} file not found: {path}")
        logger.debug(f"  ✓ {name} file exists ({path.stat().st_size} bytes)")
    
    if include_system_cas:
        # Start with default context (includes system CA bundle)
        ssl_context = ssl.create_default_context()
        logger.debug(f"  Using default context with system CAs")
        
        # Add our custom CA on top of system CAs
        logger.info("Loading custom CA certificate (in addition to system CAs)...")
        ssl_context.load_verify_locations(cafile=str(ca_path))
    else:
        # Create bare context with only our CA (like curl --cacert)
        ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        ssl_context.verify_mode = ssl.CERT_REQUIRED
        ssl_context.check_hostname = True
        logger.debug(f"  Using custom CA only (no system CAs)")
        
        logger.info("Loading CA certificate...")
        ssl_context.load_verify_locations(cafile=str(ca_path))
    
    logger.debug(f"  ✓ CA certificate loaded")
    
    # Load client certificate and key for mTLS
    logger.info("Loading client certificate and key...")
    ssl_context.load_cert_chain(certfile=str(cert_path), keyfile=str(key_path))
    logger.debug(f"  ✓ Client cert chain loaded")
    
    # Log SSL context info
    logger.debug(f"  Minimum TLS version: {ssl_context.minimum_version.name}")
    logger.debug(f"  Maximum TLS version: {ssl_context.maximum_version.name}")
    
    return ssl_context


def test_post(url: str, ssl_context: ssl.SSLContext, data: dict) -> bool:
    """
    Perform POST request with mTLS.
    """
    logger.info(f"Preparing POST request to: {url}")
    logger.debug(f"  Payload size: {len(json.dumps(data))} bytes")
    
    try:
        logger.info("Creating httpx client with SSL context...")
        with httpx.Client(
            verify=ssl_context,
            timeout=30.0
        ) as client:
            logger.info("Sending POST request...")
            
            response = client.post(
                url,
                json=data,
                headers={"Content-Type": "application/json"}
            )
            
            logger.info(f"Response received!")
            logger.info(f"  Status code: {response.status_code}")
            logger.debug(f"  Response headers:")
            for name, value in response.headers.items():
                logger.debug(f"    {name}: {value}")
            
            logger.debug(f"  Response body ({len(response.content)} bytes):")
            try:
                body = response.json()
                logger.debug(f"    {json.dumps(body, indent=2)}")
            except:
                logger.debug(f"    {response.text[:500]}")
            
            response.raise_for_status()
            logger.info("✓ POST successful!")
            return True
            
    except httpx.HTTPStatusError as e:
        logger.error(f"HTTP error: {e.response.status_code}")
        logger.error(f"  Response: {e.response.text}")
        return False
        
    except ssl.SSLError as e:
        logger.error(f"SSL error: {e}")
        logger.error(f"  SSL error code: {e.reason}")
        return False
        
    except Exception as e:
        logger.error(f"Request failed: {type(e).__name__}: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Test mTLS POST request with verbose logging",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example:
    %(prog)s https://api.example.com/endpoint \\
        --ca ~/.config/cauth/ca.pem \\
        --cert ~/.config/cauth/crt.pem \\
        --key ~/.config/cauth/key.pem
        
    # With custom JSON data
    %(prog)s https://api.example.com/endpoint \\
        --ca ca.pem --cert crt.pem --key key.pem \\
        --data '{"test": "value"}'
"""
    )
    parser.add_argument(
        "url",
        help="URL to POST to"
    )
    parser.add_argument(
        "--ca",
        required=True,
        type=Path,
        help="Path to CA certificate file (PEM)"
    )
    parser.add_argument(
        "--cert",
        required=True,
        type=Path,
        help="Path to client certificate file (PEM)"
    )
    parser.add_argument(
        "--key",
        required=True,
        type=Path,
        help="Path to client private key file (PEM)"
    )
    parser.add_argument(
        "--data", "-d",
        type=str,
        help="JSON data to POST (default: test payload)"
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Reduce logging verbosity"
    )
    parser.add_argument(
        "--no-system-cas",
        action="store_true",
        help="Don't include system CA bundle (use only --ca file)"
    )
    
    args = parser.parse_args()
    
    # Adjust logging level
    if args.quiet:
        logging.getLogger().setLevel(logging.INFO)
        logging.getLogger("httpx").setLevel(logging.WARNING)
        logging.getLogger("httpcore").setLevel(logging.WARNING)
    
    # Expand paths
    ca_path = args.ca.expanduser()
    cert_path = args.cert.expanduser()
    key_path = args.key.expanduser()
    
    logger.info("=" * 60)
    logger.info("mTLS POST Test")
    logger.info("=" * 60)
    logger.info(f"URL: {args.url}")
    logger.info(f"CA:  {ca_path}")
    logger.info(f"Cert: {cert_path}")
    logger.info(f"Key: {key_path}")
    logger.info("=" * 60)
    
    # Parse or generate test data
    if args.data:
        try:
            data = json.loads(args.data)
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON data: {e}")
            sys.exit(1)
    else:
        data = {
            "test": True,
            "timestamp": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            "message": "mTLS test from test_mtls.py"
        }
    
    logger.info(f"Payload: {json.dumps(data)}")
    logger.info("=" * 60)
    
    try:
        # Create SSL context
        ssl_context = create_ssl_context(
            ca_path,
            cert_path, 
            key_path,
            include_system_cas=not args.no_system_cas
        )
        
        # Perform POST
        success = test_post(args.url, ssl_context, data)
        
        logger.info("=" * 60)
        if success:
            logger.info("✓ TEST PASSED")
            sys.exit(0)
        else:
            logger.error("✗ TEST FAILED")
            sys.exit(1)
            
    except Exception as e:
        logger.error(f"Test failed with exception: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
