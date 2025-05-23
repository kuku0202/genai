#!/usr/bin/env python3
"""
Optimized PPT Video Generator
Generate videos from PowerPoint slides with audio narration
"""

import os
import argparse
import sys
import json
import tempfile
import shutil
import asyncio
import subprocess
from pathlib import Path
from tqdm import tqdm
import time

# Check for required libraries
try:
    import edge_tts
    has_edge_tts = True
except ImportError:
    has_edge_tts = False
    print("Warning: edge-tts not installed. Installing it is recommended for best quality.")
    print("Run: pip install edge-tts")

try:
    from gtts import gTTS
    has_gtts = True
except ImportError:
    has_gtts = False
    print("Warning: gtts not installed. This is a fallback TTS option.")
    print("Run: pip install gtts")

if not has_edge_tts and not has_gtts:
    print("Error: No TTS engine available. Please install edge-tts or gtts.")
    print("Run: pip install edge-tts")
    sys.exit(1)


class VideoGenerator:
    def __init__(self, slides_dir, scripts_dir, output_dir="output", 
                 output_video="presentation_video.mp4", 
                 tts_engine="edge", voice="en-US-ChristopherNeural",
                 fps=30, slide_duration_multiplier=1.0, resolution=(1280, 720)):
        """
        Initialize the video generator
        
        Args:
            slides_dir: Directory containing slide images
            scripts_dir: Directory containing script files
            output_dir: Directory for output files
            output_video: Output video filename
            tts_engine: TTS engine to use ("edge" or "gtts")
            voice: Voice ID for TTS
            fps: Frames per second for video
            slide_duration_multiplier: Multiplier for slide duration
            resolution: Video resolution (width, height)
        """
        self.slides_dir = Path(slides_dir)
        self.scripts_dir = Path(scripts_dir)
        self.output_dir = Path(output_dir)
        self.output_video = output_video
        self.tts_engine = tts_engine
        self.voice = voice
        self.fps = fps
        self.slide_duration_multiplier = slide_duration_multiplier
        self.resolution = resolution
        self.temp_dir = Path(tempfile.mkdtemp())
        
        # Ensure output directory exists
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Verify dependencies
        self._check_dependencies()
    
    def __del__(self):
        """Clean up temporary files"""
        try:
            shutil.rmtree(self.temp_dir)
        except:
            pass
    
    def _check_dependencies(self):
        """Check if required dependencies are installed"""
        # Check for FFmpeg
        ffmpeg_path = shutil.which('ffmpeg')
        if not ffmpeg_path:
            raise RuntimeError("FFmpeg is required but not found. Please install FFmpeg.")
        
        # Check for FFprobe
        ffprobe_path = shutil.which('ffprobe')
        if not ffprobe_path:
            raise RuntimeError("FFprobe is required but not found. It should be installed with FFmpeg.")
        
        print(f"Using FFmpeg: {ffmpeg_path}")
        print(f"Using FFprobe: {ffprobe_path}")
        
        # Check TTS engine
        if self.tts_engine == "edge" and not has_edge_tts:
            print("Warning: edge-tts not available, falling back to gtts if available")
            self.tts_engine = "gtts" if has_gtts else None
        
        if self.tts_engine == "gtts" and not has_gtts:
            print("Warning: gtts not available")
            self.tts_engine = "edge" if has_edge_tts else None
        
        if not self.tts_engine:
            raise RuntimeError("No TTS engine available. Please install edge-tts or gtts.")
        
        print(f"Using TTS engine: {self.tts_engine}")
        print(f"Using voice: {self.voice}")
    
    def collect_slides(self):
        """Collect slide images from slides directory"""
        slide_files = []
        
        # Try different naming patterns
        patterns = ["slide_*.png", "Slide*.png", "slide-*.png", "*.png"]
        
        for pattern in patterns:
            files = sorted(self.slides_dir.glob(pattern), 
                         key=lambda x: int(''.join(filter(str.isdigit, x.stem))) if any(c.isdigit() for c in x.stem) else 0)
            if files:
                slide_files = files
                break
        
        if not slide_files:
            raise RuntimeError(f"No slide images found in {self.slides_dir}")
        
        print(f"Found {len(slide_files)} slide images")
        return slide_files
    
    def find_script_for_slide(self, slide_num):
        """Find script for a specific slide number"""
        # Check different possible locations
        possible_paths = [
            self.scripts_dir / f"page_{slide_num}" / "teacher_script.txt",
            self.scripts_dir / f"page_{slide_num}_script.txt",
            self.scripts_dir / f"slide_{slide_num}_script.txt",
            self.scripts_dir / f"script_{slide_num}.txt"
        ]
        
        # Also check main output dir if different from scripts_dir
        if self.scripts_dir != self.output_dir:
            possible_paths.append(self.output_dir / f"page_{slide_num}" / "teacher_script.txt")
        
        for path in possible_paths:
            if path.exists():
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        content = f.read().strip()
                        return content
                except Exception as e:
                    print(f"Error reading script file {path}: {e}")
        
        print(f"Warning: No script found for slide {slide_num}")
        return f"Slide {slide_num}"
    
    async def generate_audio_edge_tts(self, script, output_file):
        """Generate audio using Edge TTS"""
        try:
            if not script:
                return False
                
            communicate = edge_tts.Communicate(script, self.voice)
            await communicate.save(str(output_file))
            
            return output_file.exists() and output_file.stat().st_size > 0
        except Exception as e:
            print(f"Error generating audio with Edge TTS: {e}")
            return False
    
    def generate_audio_gtts(self, script, output_file):
        """Generate audio using Google TTS"""
        try:
            if not script:
                return False
                
            tts = gTTS(text=script, lang='en' if not self.voice.startswith('zh') else 'zh-CN')
            tts.save(str(output_file))
            
            return output_file.exists() and output_file.stat().st_size > 0
        except Exception as e:
            print(f"Error generating audio with Google TTS: {e}")
            return False
    
    async def generate_audio(self, script, slide_num):
        """Generate audio for a script"""
        output_file = self.temp_dir / f"audio_{slide_num:03d}.mp3"
        
        if self.tts_engine == "edge":
            success = await self.generate_audio_edge_tts(script, output_file)
        elif self.tts_engine == "gtts":
            success = self.generate_audio_gtts(script, output_file)
        else:
            print(f"Unknown TTS engine: {self.tts_engine}")
            return None
        
        if success:
            return output_file
        
        print(f"Failed to generate audio for slide {slide_num}")
        return None
    
    def get_audio_duration(self, audio_file):
        """Get duration of an audio file in seconds"""
        if not audio_file or not audio_file.exists():
            return 5.0  # Default duration
        
        try:
            cmd = [
                "ffprobe",
                "-v", "error",
                "-show_entries", "format=duration",
                "-of", "json",
                str(audio_file)
            ]
            
            output = subprocess.check_output(cmd, stderr=subprocess.DEVNULL).decode("utf-8")
            data = json.loads(output)
            duration = float(data["format"]["duration"])
            
            # Apply duration multiplier
            duration *= self.slide_duration_multiplier
            
            # Ensure minimum duration
            return max(duration, 3.0)
        except Exception as e:
            print(f"Error getting audio duration: {e}")
            return 5.0
    
    def create_slide_video(self, slide_file, audio_file, slide_num):
        """Create video for a single slide with audio"""
        if not slide_file.exists():
            print(f"Error: Slide file not found: {slide_file}")
            return None
        
        # Get audio duration
        duration = self.get_audio_duration(audio_file)
        
        # Create output video file
        video_file = self.temp_dir / f"video_{slide_num:03d}.mp4"
        
        # Build FFmpeg command
        width, height = self.resolution
        cmd = [
            "ffmpeg", "-y",
            "-loop", "1",
            "-i", str(slide_file),
            "-t", str(duration)
        ]
        
        # Add audio if available
        if audio_file and audio_file.exists():
            cmd.extend([
                "-i", str(audio_file),
                "-c:a", "aac",
                "-b:a", "192k"
            ])
        
        # Add video settings
        cmd.extend([
            "-c:v", "libx264",
            "-tune", "stillimage",
            "-pix_fmt", "yuv420p",
            "-vf", f"scale={width}:{height}:force_original_aspect_ratio=decrease,pad={width}:{height}:(ow-iw)/2:(oh-ih)/2",
            "-shortest",
            str(video_file)
        ])
        
        try:
            # Run FFmpeg
            subprocess.run(cmd, check=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
            
            if video_file.exists():
                return video_file
            else:
                print(f"Error: Failed to create video for slide {slide_num}")
                return None
        except subprocess.CalledProcessError as e:
            print(f"Error creating video for slide {slide_num}:")
            print(e.stderr.decode("utf-8"))
            return None
    
    def concatenate_videos(self, video_files):
        """Concatenate all videos into a single file"""
        if not video_files:
            print("No videos to concatenate")
            return None
        
        # Create a list file for FFmpeg
        concat_file = self.temp_dir / "concat_list.txt"
        with open(concat_file, "w", encoding="utf-8") as f:
            for video in video_files:
                f.write(f"file '{video.absolute()}'\n")
        
        # Output file
        output_file = self.output_dir / self.output_video
        
        # Concatenate using FFmpeg
        cmd = [
            "ffmpeg", "-y",
            "-f", "concat",
            "-safe", "0",
            "-i", str(concat_file),
            "-c", "copy",
            str(output_file)
        ]
        
        try:
            subprocess.run(cmd, check=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
            
            if output_file.exists():
                return output_file
        except subprocess.CalledProcessError as e:
            print(f"Error concatenating videos:")
            print(e.stderr.decode("utf-8"))
        
        return None
    
    async def generate_video(self):
        """Generate the complete presentation video"""
        start_time = time.time()
        print("Starting video generation process...")
        
        # Step 1: Collect slides
        slide_files = self.collect_slides()
        
        # Step 2: Process each slide
        video_files = []
        
        for i, slide_file in enumerate(tqdm(slide_files, desc="Processing slides")):
            slide_num = i + 1
            
            # Get script for this slide
            script = self.find_script_for_slide(slide_num)
            
            # Generate audio
            audio_file = await self.generate_audio(script, slide_num)
            
            # Create video
            video_file = self.create_slide_video(slide_file, audio_file, slide_num)
            
            if video_file:
                video_files.append(video_file)
        
        # Step 3: Concatenate videos
        if video_files:
            print(f"Concatenating {len(video_files)} video segments...")
            final_video = self.concatenate_videos(video_files)
            
            if final_video:
                duration = time.time() - start_time
                print(f"Video generation completed in {duration:.2f} seconds")
                print(f"Output video: {final_video}")
                return final_video
        
        print("Video generation failed")
        return None


async def list_edge_tts_voices():
    """List available Edge TTS voices"""
    try:
        voices = await edge_tts.list_voices()
        
        print("Available Edge TTS voices:")
        print("%-30s %s" % ("Voice Name", "Gender"))
        print("-" * 50)
        
        for voice in sorted(voices, key=lambda v: (v["Locale"], v["ShortName"])):
            print("%-30s %s" % (voice["ShortName"], voice["Gender"]))
        
        print("\nExample usage:")
        print("English (US):   en-US-ChristopherNeural (Male)")
        print("English (US):   en-US-JennyNeural (Female)")
        print("Chinese:        zh-CN-YunxiNeural (Male)")
        print("Chinese:        zh-CN-XiaoxiaoNeural (Female)")
    except Exception as e:
        print(f"Error listing voices: {e}")


async def main():
    parser = argparse.ArgumentParser(description="Generate a presentation video from slides and scripts")
    parser.add_argument("--slides_dir", default="slides", help="Directory containing slide images")
    parser.add_argument("--scripts_dir", default="output", help="Directory containing script files")
    parser.add_argument("--output_dir", default=".", help="Output directory")
    parser.add_argument("--output_video", default="presentation_video.mp4", help="Output video filename")
    parser.add_argument("--tts", default="edge", choices=["edge", "gtts"], help="TTS engine to use")
    parser.add_argument("--voice", default="en-US-ChristopherNeural", help="Voice ID for TTS")
    parser.add_argument("--fps", type=int, default=30, help="Frames per second")
    parser.add_argument("--duration_multiplier", type=float, default=1.0, 
                        help="Multiplier for slide duration (e.g., 1.2 for 20% longer)")
    parser.add_argument("--width", type=int, default=1280, help="Video width")
    parser.add_argument("--height", type=int, default=720, help="Video height")
    parser.add_argument("--list_voices", action="store_true", help="List available Edge TTS voices and exit")
    
    args = parser.parse_args()
    
    # List voices if requested
    if args.list_voices:
        if has_edge_tts:
            await list_edge_tts_voices()
        else:
            print("Edge TTS is not installed. Install with: pip install edge-tts")
        return
    
    # Create and run generator
    generator = VideoGenerator(
        slides_dir=args.slides_dir,
        scripts_dir=args.scripts_dir,
        output_dir=args.output_dir,
        output_video=args.output_video,
        tts_engine=args.tts,
        voice=args.voice,
        fps=args.fps,
        slide_duration_multiplier=args.duration_multiplier,
        resolution=(args.width, args.height)
    )
    
    await generator.generate_video()


if __name__ == "__main__":
    asyncio.run(main())