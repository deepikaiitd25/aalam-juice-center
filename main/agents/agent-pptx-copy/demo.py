#!/usr/bin/env python3
"""
Demo script for PPTX agent with image generation.
Shows how to create presentations with AI-generated images.
"""

import os
import sys
import argparse
from dotenv import load_dotenv

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from pptx_toolset import PptxToolset

load_dotenv()

def demo_with_images():
    """Demo presentation with AI-generated images."""
    print("🎨 Creating presentation with AI-generated images...")

    toolset = PptxToolset(host="localhost", port=8000)

    # Generate images for the presentation
    print("Generating images...")

    # Fetch stock photos
    title_image = toolset.fetch_stock_photo("business presentation background abstract geometric patterns")

    # Generate charts using matplotlib
    chart_image = toolset.generate_chart(
        chart_type="line",
        title="Business Growth Q1-Q4 2024",
        labels=["Q1", "Q2", "Q3", "Q4"],
        values=[120, 145, 168, 195]
    )

    team_image = toolset.fetch_stock_photo("professional team collaboration office meeting")

    print(f"Title image: {title_image.status}")
    print(f"Chart image: {chart_image.status}")
    print(f"Team image: {team_image.status}")

    # Create presentation slides
    slides_data = [
        {
            "type": "title",
            "title": "Company Growth Strategy 2024",
            "subtitle": "Driving Innovation and Sustainable Development",
            "image_url": title_image.image_url if title_image.status == "success" else None
        },
        {
            "type": "content",
            "title": "Market Performance",
            "bullets": [
                "35% YoY revenue growth achieved",
                "Expanded market share by 12%",
                "New customer acquisition up 45%",
                "Improved customer satisfaction scores"
            ],
            "image_url": chart_image.image_url if chart_image.status == "success" else None
        },
        {
            "type": "content",
            "title": "Team Excellence",
            "bullets": [
                "Cross-functional collaboration enhanced",
                "Innovation pipeline strengthened",
                "Professional development programs launched",
                "Diverse talent acquisition improved"
            ],
            "image_url": team_image.image_url if team_image.status == "success" else None
        },
        {
            "type": "two_column",
            "title": "Strategic Initiatives",
            "left_title": "Q1 Achievements",
            "left_bullets": [
                "Digital transformation completed",
                "New product line launched",
                "Process optimization implemented"
            ],
            "right_title": "Q2 Priorities",
            "right_bullets": [
                "Market expansion planning",
                "Technology infrastructure upgrade",
                "Sustainability goals alignment"
            ]
        },
        {
            "type": "closing",
            "title": "Thank You",
            "subtitle": "Questions & Discussion"
        }
    ]

    result = toolset.generate_pptx("company_strategy_2024", slides_data, "blue")

    if result.status == "success":
        print("✅ Presentation created successfully!")
        print(f"📁 File: {result.file_url}")
        print("\n📊 Presentation Summary:")
        print("- Title slide with custom background")
        print("- 2 content slides with infographics")
        print("- 1 comparison slide (two-column)")
        print("- Closing slide")
        print(f"- Images generated: {sum(1 for img in [title_image, chart_image, team_image] if img.status == 'success')}/3")
    else:
        print(f"❌ Failed to create presentation: {result.error_message}")

def demo_without_images():
    """Demo presentation without images (fallback mode)."""
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

    result = toolset.generate_pptx("project_status_q1", slides_data, "green")

    if result.status == "success":
        print("✅ Presentation created successfully!")
        print(f"📁 File: {result.file_url}")
    else:
        print(f"❌ Failed to create presentation: {result.error_message}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PPTX Agent Demo")
    parser.add_argument("--with-images", action="store_true",
                       help="Generate presentation with AI images (requires OpenAI API key)")
    parser.add_argument("--without-images", action="store_true",
                       help="Generate presentation without images")

    args = parser.parse_args()

    if args.with_images:
        demo_with_images()
    elif args.without_images:
        demo_without_images()
    else:
        print("PPTX Agent Demo")
        print("Usage:")
        print("  python3 demo.py --with-images    # Create presentation with AI-generated images")
        print("  python3 demo.py --without-images # Create presentation without images")
        print("\nNote: --with-images requires OPENAI_API_KEY in .env file")
