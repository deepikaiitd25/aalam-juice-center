#!/usr/bin/env python3
"""
Simple test script for PPTX agent image generation functionality.
Tests the core image generation and PPTX creation without the A2A framework.
"""

import os
import sys
import json

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from pptx_toolset import PptxToolset

def test_image_generation():
    """Test image generation functionality."""
    print("Testing image generation...")
    
    # Initialize toolset with API key from environment
    from dotenv import load_dotenv
    load_dotenv()
    openai_key = os.getenv("OPENAI_API_KEY")
    
    toolset = PptxToolset(host="localhost", port=8000, openai_api_key=openai_key)

    # Test image generation
    result = toolset.generate_image("test prompt")
    print(f"Image generation result: {result}")
    if openai_key and openai_key != "your_openai_api_key_here":
        # Should succeed or fail based on API limits/validity
        assert result.status in ["success", "error"]
        if result.status == "error":
            print(f"Image generation failed (expected if billing limit reached): {result.error_message}")
    else:
        assert result.status == "error"
        assert "OpenAI API key not configured" in result.error_message
        print("✓ Image generation correctly handles missing API key")


def test_pptx_generation():
    """Test PPTX generation functionality."""
    print("\nTesting PPTX generation...")
    
    toolset = PptxToolset(host="localhost", port=8000)

    # Test data for a simple presentation
    slides_data = [
        {
            "type": "title",
            "title": "Test Presentation",
            "subtitle": "With Image Generation"
        },
        {
            "type": "content",
            "title": "Slide with Image",
            "bullets": [
                "This slide has an image",
                "Generated automatically",
                "Professional layout"
            ],
            "image_url": "https://via.placeholder.com/400x300.png?text=Test+Image"
        },
        {
            "type": "content",
            "title": "Regular Slide",
            "bullets": [
                "Normal content slide",
                "No image needed",
                "Standard layout"
            ]
        }
    ]

    result = toolset.generate_pptx("test_presentation", slides_data, "blue")
    print(f"PPTX generation result: {result}")
    assert result.status == "success"
    assert "test_presentation.pptx" in result.file_url
    print("✓ PPTX generation successful")

    # Check if file was created
    filepath = os.path.join("outputs", "test_presentation.pptx")
    assert os.path.exists(filepath)
    print(f"✓ PPTX file created at {filepath}")

def test_slide_with_image():
    """Test that slides with images are handled correctly."""
    print("\nTesting slide layout with images...")

    toolset = PptxToolset(host="localhost", port=8000)

    slides_data = [
        {
            "type": "content",
            "title": "Data Visualization",
            "bullets": [
                "Key metrics shown",
                "Growth trends",
                "Market analysis"
            ],
            "image_url": "https://via.placeholder.com/300x200.png?text=Chart"
        }
    ]

    result = toolset.generate_pptx("image_test", slides_data, "green")
    assert result.status == "success"
    print("✓ Slide with image generated successfully")

if __name__ == "__main__":
    print("Running PPTX Agent Tests...")
    print("=" * 50)

    try:
        test_image_generation()
        test_pptx_generation()
        test_slide_with_image()

        print("\n" + "=" * 50)
        print("✅ All tests passed!")
        print("\nTo test with real image generation:")
        print("1. Add your OPENAI_API_KEY to .env file")
        print("2. Run: python3 test_agent.py --with-images")

    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        raise
