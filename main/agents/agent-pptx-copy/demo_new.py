#!/usr/bin/env python3
"""
Demo script for PPTX agent v2 - with matplotlib charts and stock photos.
"""

import os
import sys
import argparse
import asyncio
from dotenv import load_dotenv

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from pptx_toolset import PptxToolset

load_dotenv()

async def demo_with_images():
    """Demo presentation with charts and stock photos."""
    print("🎨 Creating presentation with charts and stock photos...")

    toolset = PptxToolset(host="localhost", port=8000)

    # Create presentation slides with matplotlib charts and stock photos
    slides_data = [
        {
            "type": "title",
            "title": "Company Growth Strategy 2024",
            "subtitle": "Driving Innovation and Sustainable Development"
        },
        {
            "type": "image",
            "keyword": "business financial dashboard data analytics report",
            "fallback_keywords": ["corporate profit growth chart", "revenue metrics dashboard", "financial statistics office"],
            "title": "Market Overview",
            "body": "Strategic market analysis and performance metrics"
        },
        {
            "type": "chart",
            "chart_type": "line",
            "title": "Business Growth Q1-Q4 2024",
            "labels": ["Q1", "Q2", "Q3", "Q4"],
            "values": [120, 145, 168, 195]
        },
        {
            "type": "content",
            "title": "Key Achievements",
            "bullets": [
                "35% YoY revenue growth achieved",
                "Expanded market share by 12%",
                "New customer acquisition up 45%",
                "Improved customer satisfaction scores"
            ]
        },
        {
            "type": "chart",
            "chart_type": "bar",
            "title": "Revenue by Region",
            "labels": ["North", "South", "East", "West"],
            "values": [250, 180, 220, 200]
        },
        {
            "type": "two_column",
            "title": "Strategic Initiatives vs Achievements",
            "left_title": "Planned for Q1",
            "left_bullets": [
                "Digital transformation",
                "New product launch",
                "Process optimization"
            ],
            "right_title": "Completed in Q1",
            "right_bullets": [
                "✓ Digital transformation completed",
                "✓ New product line launched",
                "✓ Process optimization implemented"
            ]
        },
        {
            "type": "closing",
            "title": "Thank You",
            "subtitle": "Questions & Discussion"
        }
    ]

    result = await toolset.generate_pptx("company_strategy_2024", slides_data, "blue")
    print(result)
    print("\n✅ Test completed!")


async def demo_without_images():
    """Demo presentation without images."""
    print("📊 Creating presentation without images...")

    toolset = PptxToolset(host="localhost", port=8000)

    slides_data = [
        {
            "type": "title",
            "title": "Project Status Report",
            "subtitle": "Q1 2024 Achievements"
        },
        {
            "type": "content",
            "title": "Key Accomplishments",
            "bullets": [
                "Successfully launched 3 new features",
                "Improved system performance by 40%",
                "Reduced customer support tickets by 25%",
                "Enhanced security protocols implemented"
            ]
        },
        {
            "type": "two_column",
            "title": "Current vs Target Metrics",
            "left_title": "Current Performance",
            "left_bullets": ["Revenue: $2.1M", "Users: 15,000", "Satisfaction: 4.2/5"],
            "right_title": "Q2 Targets",
            "right_bullets": ["Revenue: $2.8M", "Users: 20,000", "Satisfaction: 4.5/5"]
        },
        {
            "type": "closing",
            "title": "Next Steps",
            "subtitle": "Continued focus on growth and innovation"
        }
    ]

    result = await toolset.generate_pptx("project_status_q1", slides_data, "green")
    print(result)
    print("\n✅ Test completed!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PPTX Agent v2 Demo")
    parser.add_argument("--with-images", action="store_true",
                       help="Generate presentation with charts and stock photos")
    parser.add_argument("--without-images", action="store_true",
                       help="Generate presentation without images")

    args = parser.parse_args()

    if args.with_images:
        asyncio.run(demo_with_images())
    elif args.without_images:
        asyncio.run(demo_without_images())
    else:
        print("PPTX Agent v2 Demo")
        print("Usage:")
        print("  python3 demo_new.py --with-images    # Create presentation with charts & stock photos")
        print("  python3 demo_new.py --without-images # Create presentation without images")
