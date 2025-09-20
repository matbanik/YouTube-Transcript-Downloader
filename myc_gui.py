import os
import time
import json
import random
import re
import glob
import shutil
import zipfile
import threading
import queue
from datetime import datetime, timedelta
import hashlib

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, scrolledtext

# --- Dependency Checks ---
try:
    from youtube_transcript_api import YouTubeTranscriptApi
    from youtube_transcript_api._errors import TranscriptsDisabled, NoTranscriptFound
    import yt_dlp
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib.colors import blue, darkblue, mediumseagreen, darkgreen
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.lib.enums import TA_LEFT, TA_CENTER
except ImportError as e:
    # Use a basic Tkinter window for the error if the main app can't start
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing Dependencies",
        f"A required library is missing: {e.name}.\n\nPlease run 'pip install -r requirements.txt' to install all dependencies."
    )
    exit()

# --- Configuration Management ---
CONFIG_FILE = 'download_config.json'
DEFAULT_CONFIG = {
    'channels': [
        {"url": "https://www.youtube.com/@drekberg", "completed": False},
        {"url": "https://www.youtube.com/c/paulsaladinomd", "completed": False}
    ],
    'requests_per_hour': 300,
    'delay_between_requests': 6,
    'max_retries': 3,
    'retry_delay': 30,
    'batch_size': 50,
    'playlist_end': 0,
    'pause_between_batches': 300,
    'max_words_per_pdf': 490000
}
COMPLETED_PREFIX = "✅ [DONE] "

def load_config():
    """Load configuration from file, create with defaults, or migrate old format."""
    config_to_save = False
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
            # Migrate old string-based channel list to new object-based list
            if config.get('channels') and isinstance(config['channels'][0], str):
                config['channels'] = [{"url": url, "completed": False} for url in config['channels']]
                config_to_save = True
        except (json.JSONDecodeError, IndexError):
            config = DEFAULT_CONFIG.copy()
            config_to_save = True
    else:
        config = DEFAULT_CONFIG.copy()
        config_to_save = True

    # Ensure all default keys exist
    for key, value in DEFAULT_CONFIG.items():
        if key not in config:
            config[key] = value
            config_to_save = True

    if 'channels' not in config or not isinstance(config['channels'], list):
        config['channels'] = []
        config_to_save = True

    if config_to_save:
        save_config(config)

    return config


def save_config(config):
    """Save configuration to file."""
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=4)

# --- Backend Processing Logic ---
class TranscriptProcessor:
    def __init__(self, config, log_callback, is_paused_event, is_running_event, completion_callback):
        self.config = config
        self.log = log_callback
        self.is_paused = is_paused_event
        self.is_running = is_running_event
        self.completion_callback = completion_callback
        self.pause_logged = False

    def run(self):
        self.is_running.set()
        self.log(("INFO", "🚀 Starting Multi-Channel YouTube Transcript Downloader"))

        channels_to_process = list(self.config['channels'])

        for channel_info in channels_to_process:
            if not self.is_running.is_set():
                self.log(("WARNING", "🛑 Processing stopped by user."))
                break
            
            self.check_pause()

            channel_url = channel_info.get("url")
            channel_name = self.extract_channel_name(channel_url)

            if channel_info.get("completed"):
                self.log(("INFO", f"➡️ Skipping already completed channel: {channel_name}"))
                continue
            
            try:
                result = self.process_channel(channel_url, channel_name)
                if result.get('cleanup_complete'):
                    self.log(("SUCCESS", f"✅ Channel {channel_name} fully processed and cleaned up."))
                    self.completion_callback(channel_url) # Signal completion to UI
            except Exception as e:
                self.log(("ERROR", f"❌ Unexpected fatal error processing {channel_url}: {e}"))

        self.log(("INFO", "\n🎉 Run finished."))
        self.is_running.clear()

    def check_pause(self):
        """Check if processing is paused and wait until resumed."""
        if self.is_paused.is_set() and not self.pause_logged:
            self.log(("INFO", "⏸️ Processing paused. Press 'Resume' to continue."))
            self.pause_logged = True
        
        while self.is_paused.is_set():
            time.sleep(1)
            if not self.is_running.is_set():
                break
        
        if self.pause_logged and not self.is_paused.is_set() and self.is_running.is_set():
            self.log(("INFO", "▶️ Processing resumed."))
            self.pause_logged = False

    def generate_channel_id(self, channel_url):
        normalized_url = channel_url.strip().lower().rstrip('/')
        return hashlib.md5(normalized_url.encode()).hexdigest()[:12]

    def extract_channel_name(self, channel_url):
        url = channel_url.rstrip('/')
        patterns = {'/@': '/@', '/c/': '/c/', '/channel/': '/channel/', '/user/': '/user/'}
        for key, split_by in patterns.items():
            if key in url:
                name = url.split(split_by)[-1]
                break
        else:
            name = url.split('/')[-1]
        name = re.sub(r'[<>:"/\\|?*]', '', name)
        return re.sub(r'[^\w\-_.]', '_', name)

    def process_channel(self, channel_url, channel_name):
        if not self.is_running.is_set(): return {}
        self.log(("INFO", f"\n{'='*60}\n🎬 Processing Channel: {channel_name}\n{'='*60}"))
        self.check_pause()
        channel_id = self.generate_channel_id(channel_url)
        progress_file = f'progress_{channel_id}.json'
        output_dir = f'transcripts_{channel_name}_{channel_id}'
        progress = self.load_progress(progress_file)
        progress.update({'channel_url': channel_url, 'channel_name': channel_name, 'channel_id': channel_id})
        video_infos = self.get_video_ids_from_channel(channel_url)
        if not video_infos:
            self.log(("ERROR", f"❌ No videos found for {channel_name}."))
            return {'success': False, 'channel': channel_name, 'error': 'No videos found'}
        existing_ids = self.get_existing_transcripts(output_dir)
        progress['completed_videos'] = list(set(progress.get('completed_videos', [])) | existing_ids)
        remaining_videos = [v for v in video_infos if v['id'] not in progress['completed_videos']]
        if not remaining_videos:
            self.log(("INFO", f"✅ All {len(video_infos)} videos for {channel_name} already downloaded."))
        else:
            self.log(("INFO", f"📋 Processing {len(remaining_videos)} new videos for {channel_name}..."))
            for i, video_info in enumerate(remaining_videos):
                if not self.is_running.is_set():
                    self.save_progress(progress, progress_file)
                    return {}
                self.check_pause()
                video_id, video_title = video_info['id'], video_info['title']
                self.log(("INFO", f"\n[{i+1}/{len(remaining_videos)}] ID: {video_id} - {video_title[:50]}..."))
                success = self.download_single_transcript(video_id, video_title, output_dir, progress)
                if not success:
                    self.log(("ERROR", f"❌ Failed to download transcript for {video_id} after retries."))
                if (i + 1) % 10 == 0:
                    self.save_progress(progress, progress_file)
            self.save_progress(progress, progress_file)
            self.log(("INFO", f"\n📊 Download phase complete for {channel_name}."))
        
        if not self.is_running.is_set(): return {}

        pdf_dir = f'pdf_{channel_name}_{channel_id}'
        if self.create_pdfs_from_transcripts(output_dir, pdf_dir, channel_name):
            if self.perform_cleanup(channel_name, channel_id, output_dir, progress_file, pdf_dir):
                return {'success': True, 'channel': channel_name, 'cleanup_complete': True}
        return {'success': True, 'channel': channel_name, 'cleanup_complete': False}

    def download_single_transcript(self, video_id, video_title, output_dir, progress):
        retries = 0
        while retries < self.config['max_retries']:
            self.check_pause()
            if not self.is_running.is_set(): return False
            try:
                # --- MODIFICATION: Handle 'random' delay ---
                delay_setting = self.config.get('delay_between_requests', 6)
                if str(delay_setting).lower() == 'random':
                    delay = random.randint(1, 77)
                    self.log(("INFO", f"  -> Using random delay: {delay}s"))
                else:
                    # Use the original logic for numeric delays
                    delay = float(delay_setting) * random.uniform(0.8, 1.2)
                
                time.sleep(delay)

                transcript = YouTubeTranscriptApi.get_transcript(video_id)
                self.save_transcript(video_id, video_title, transcript, output_dir)
                self.log(("SUCCESS", f"  -> Transcript saved for {video_id}"))
                progress.setdefault('completed_videos', []).append(video_id)
                return True
            except (TranscriptsDisabled, NoTranscriptFound) as e:
                self.log(("WARNING", f"  -> No transcript for {video_id}: {e}"))
                progress.setdefault('failed_videos', []).append({'id': video_id, 'reason': str(e)})
                return True # Treat as success to not retry
            except Exception as e:
                error_msg = str(e)
                self.log(("ERROR", f"  -> Error downloading {video_id}: {error_msg}"))
                if "too many requests" in error_msg.lower() or "429" in error_msg:
                    self.log(("WARNING", "RATE LIMIT HIT. Consider pausing and waiting, or increasing delays."))
                retries += 1
                if retries < self.config['max_retries']:
                    self.log(("INFO", f"  -> Retrying ({retries}/{self.config['max_retries']})..."))
                    self.wait_with_countdown(self.config['retry_delay'], "Waiting before retry")
        return False

    def get_video_ids_from_channel(self, channel_url):
        self.log(("INFO", f"🔍 Extracting video IDs from: {channel_url}"))
        
        # --- MODIFICATION: Disable caching to prevent tracking ---
        ydl_opts = {
            'extract_flat': True,
            'quiet': False,
            'no_warnings': False,
            'ignoreerrors': False,
            'socket_timeout': 30,
            'logtostderr': True,
            'cachedir': False, # Prevent caching to avoid leaving traces
        }
        self.log(("INFO", "  -> yt-dlp cache disabled to prevent tracking."))

        try:
            playlist_limit = int(self.config.get('playlist_end', 0))
            if playlist_limit > 0:
                self.log(("INFO", f"  -> Applying user-defined limit: will stop after finding {playlist_limit} videos."))
                ydl_opts['playlistend'] = playlist_limit
            else:
                self.log(("INFO", "  -> No video scan limit set (playlist_end = 0). Scanning all pages."))
        except (ValueError, TypeError):
            self.log(("WARNING", "  -> Could not read 'playlist_end' setting. Scanning all pages."))
        
        video_infos = []
        self.log(("INFO", "  -> Instantiating yt-dlp... (Check terminal for verbose output)"))
        
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                self.log(("INFO", "  -> yt-dlp instantiated. Attempting to extract video info..."))
                result = ydl.extract_info(channel_url + '/videos', download=False)
                self.log(("INFO", "  -> Info extraction call completed."))
                
                if result and result.get("entries"):
                    video_infos = [{'id': e['id'], 'title': e.get('title', 'No Title')} 
                                   for e in result.get("entries", []) if e and e.get('id')]
                    self.log(("SUCCESS", f"  -> Successfully extracted {len(video_infos)} video entries."))
                else:
                    self.log(("WARNING", "  -> yt-dlp completed but found no video 'entries' in the result."))

        except Exception as e:
            self.log(("ERROR", f"❌ An error occurred during video ID extraction with yt-dlp."))
            self.log(("ERROR", f"  -> Error Type: {type(e).__name__}"))
            self.log(("ERROR", f"  -> Details: {e}"))
            self.log(("WARNING", "  -> This could be due to an outdated yt-dlp library, a network issue, or a change in YouTube's site."))
            self.log(("WARNING", "  -> Try running 'pip install --upgrade yt-dlp' in your terminal."))
        
        self.log(("INFO", f"🎯 Found {len(video_infos)} video IDs."))
        return video_infos

    def save_transcript(self, video_id, video_title, transcript_data, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        clean_title = re.sub(r'[<>:"/\\|?*]', '', video_title)[:150]
        filepath = os.path.join(output_dir, f"{clean_title} ({video_id}).txt")
        with open(filepath, "w", encoding="utf-8") as f:
            f.write('\n'.join([line['text'] for line in transcript_data]))

    def create_pdfs_from_transcripts(self, transcript_folder, pdf_folder, channel_name):
        self.log(("INFO", f"\n📑 Merging transcripts for '{channel_name}' into PDFs..."))
        os.makedirs(pdf_folder, exist_ok=True)
        txt_files = sorted(glob.glob(os.path.join(transcript_folder, "*.txt")))
        if not txt_files:
            self.log(("INFO", f"ℹ️ No new transcript files in '{transcript_folder}' to merge."))
            return True
        base_filename = os.path.join(pdf_folder, channel_name)
        pdf_number, word_count, files_in_pdf = 1, 0, 0
        doc, story, styles = self.create_pdf_document(f"{base_filename}_{pdf_number}.pdf")
        story.append(Paragraph(f"Transcripts for {channel_name} - Part {pdf_number}", styles['title']))
        for txt_file in txt_files:
            with open(txt_file, 'r', encoding='utf-8') as f: content = f.read().strip()
            file_word_count = len(content.split())
            if word_count + file_word_count > self.config['max_words_per_pdf'] and files_in_pdf > 0:
                self.log(("INFO", f"📦 Building PDF {pdf_number} ({word_count:,} words)..."))
                doc.build(story)
                pdf_number += 1; word_count, files_in_pdf = 0, 0
                doc, story, styles = self.create_pdf_document(f"{base_filename}_{pdf_number}.pdf")
                story.append(Paragraph(f"Transcripts for {channel_name} - Part {pdf_number}", styles['title']))
            story.extend([PageBreak(), Paragraph(os.path.basename(txt_file).rsplit('.', 1)[0], styles['filename']), Paragraph(content.replace('&', '&').replace('<', '<'), styles['content'])])
            word_count += file_word_count; files_in_pdf += 1
        if files_in_pdf > 0:
            self.log(("INFO", f"📦 Building final PDF {pdf_number}..."))
            doc.build(story)
        self.log(("SUCCESS", f"✅ PDFs created in '{pdf_folder}'"))
        return True

    def create_pdf_document(self, filename):
        doc = SimpleDocTemplate(filename, pagesize=letter, leftMargin=.75*inch, rightMargin=.75*inch, topMargin=1*inch, bottomMargin=1*inch)
        styles = {
            'title': ParagraphStyle('Title', fontName='Helvetica-Bold', fontSize=16, alignment=TA_CENTER, spaceAfter=20, textColor=darkblue),
            'filename': ParagraphStyle('FileName', fontName='Helvetica-Bold', fontSize=12, alignment=TA_LEFT, spaceBefore=15, spaceAfter=10, textColor=blue),
            'content': ParagraphStyle('Content', fontName='Helvetica', fontSize=10, leading=12, alignment=TA_LEFT)
        }
        return doc, [], styles

    def perform_cleanup(self, channel_name, channel_id, transcript_dir, progress_file, pdf_dir):
        self.log(("INFO", f"\n🧹 Performing cleanup for {channel_name}..."))
        archive_name = os.path.join(pdf_dir, f"archive_{channel_name}_{channel_id}.zip")
        try:
            with zipfile.ZipFile(archive_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                if os.path.exists(transcript_dir):
                    for root, _, files in os.walk(transcript_dir):
                        for file in files: zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), transcript_dir))
                    self.log(("INFO", f"  -> Archived '{transcript_dir}'"))
                if os.path.exists(progress_file):
                    zipf.write(progress_file, os.path.basename(progress_file))
                    self.log(("INFO", f"  -> Archived '{progress_file}'"))
            if os.path.exists(transcript_dir): shutil.rmtree(transcript_dir)
            if os.path.exists(progress_file): os.remove(progress_file)
            self.log(("SUCCESS", f"  -> Deleted original files and folder."))
            return True
        except Exception as e:
            self.log(("ERROR", f"❌ Cleanup failed: {e}"))
            return False

    def load_progress(self, progress_file):
        if os.path.exists(progress_file):
            with open(progress_file, 'r') as f: return json.load(f)
        return {}

    def save_progress(self, progress, progress_file):
        with open(progress_file, 'w') as f: json.dump(progress, f, indent=2)

    def get_existing_transcripts(self, output_dir):
        if not os.path.exists(output_dir): return set()
        return {m.group(1) for f in os.listdir(output_dir) if (m := re.search(r'\((\w{11})\)\.txt$', f))}

    def wait_with_countdown(self, seconds, message="Waiting"):
        for i in range(seconds, 0, -1):
            if not self.is_running.is_set(): break
            self.check_pause()
            self.log(("PROGRESS", f"⏳ {message}... {i}s remaining"))
            time.sleep(1)
        self.log(("PROGRESS", " " * 50)) # Clear line

# --- Tkinter GUI Application ---
class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("YouTube Transcript Downloader")
        self.master.geometry("800x700")
        self.pack(fill=tk.BOTH, expand=True)
        self.log_queue = queue.Queue()
        self.worker_thread = None
        self.is_paused = threading.Event()
        self.is_running = threading.Event()
        self.create_widgets()
        self.load_config_to_ui()
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.process_log_queue()
        self.update_button_states()

    def create_widgets(self):
        top_settings_frame = ttk.Frame(self, padding="10")
        top_settings_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 0))
        controls_frame = ttk.Frame(self, padding=(10, 5, 10, 10))
        controls_frame.pack(side=tk.TOP, fill=tk.X)
        console_frame = ttk.Labelframe(self, text="Console Log", padding="5")
        console_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        top_pane = ttk.PanedWindow(top_settings_frame, orient=tk.HORIZONTAL)
        top_pane.pack(fill=tk.BOTH, expand=True)
        channels_frame = ttk.Labelframe(top_pane, text="YouTube Channels", padding="5")
        top_pane.add(channels_frame, weight=1)
        config_frame = ttk.Labelframe(top_pane, text="Configuration", padding="5")
        top_pane.add(config_frame, weight=1)
        self.channel_listbox = tk.Listbox(channels_frame, height=10, selectbackground="#0078D7")
        self.channel_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.channel_buttons = []
        channel_buttons_frame = ttk.Frame(channels_frame)
        channel_buttons_frame.pack(fill=tk.X)
        btn_up = ttk.Button(channel_buttons_frame, text="Move Up", command=self.move_channel_up)
        btn_up.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        btn_down = ttk.Button(channel_buttons_frame, text="Move Down", command=self.move_channel_down)
        btn_down.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        btn_add = ttk.Button(channel_buttons_frame, text="Add", command=self.add_channel)
        btn_add.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        btn_remove = ttk.Button(channel_buttons_frame, text="Remove", command=self.remove_channel)
        btn_remove.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.channel_buttons.extend([btn_up, btn_down, btn_add, btn_remove])
        
        self.config_vars = {}
        self.config_entries = []
        for key, value in DEFAULT_CONFIG.items():
            if key == 'channels': continue
            frame = ttk.Frame(config_frame)
            frame.pack(fill=tk.X, pady=2)
            ttk.Label(frame, text=f"{key.replace('_', ' ').title()}:").pack(side=tk.LEFT, anchor='w')
            var = tk.StringVar(value=str(value))
            self.config_vars[key] = var
            entry = ttk.Entry(frame, textvariable=var)
            entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
            self.config_entries.append(entry)

        self.start_button = ttk.Button(controls_frame, text="Start Processing", command=self.start_processing)
        self.start_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.pause_button = ttk.Button(controls_frame, text="Pause", command=self.pause_processing)
        self.pause_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.stop_button = ttk.Button(controls_frame, text="Stop", command=self.stop_processing)
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.save_button = ttk.Button(controls_frame, text="Save Configuration", command=lambda: self.save_config_from_ui(silent=False))
        self.save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.console = scrolledtext.ScrolledText(console_frame, state='disabled', wrap=tk.WORD, bg="#2b2b2b", fg="#d3d3d3", insertbackground='white')
        self.console.pack(fill=tk.BOTH, expand=True)
        self.console.tag_config("SUCCESS", foreground="#A9D1A9")
        self.console.tag_config("ERROR", foreground="#FF7B7B")
        self.console.tag_config("WARNING", foreground="#FFD700")
        self.console.tag_config("INFO", foreground="#d3d3d3")

    def load_config_to_ui(self):
        self.config = load_config()
        self.channel_listbox.delete(0, tk.END)
        for i, channel_info in enumerate(self.config.get('channels', [])):
            url = channel_info.get("url")
            is_completed = channel_info.get("completed", False)
            display_text = f"{COMPLETED_PREFIX}{url}" if is_completed else url
            self.channel_listbox.insert(tk.END, display_text)
            if is_completed:
                self.channel_listbox.itemconfig(i, {'bg': '#2E4034', 'fg': '#A9D1A9'})
        for key, var in self.config_vars.items():
            var.set(str(self.config.get(key, DEFAULT_CONFIG[key])))

    def save_config_from_ui(self, silent=True):
        new_channels_config = []
        for i in range(self.channel_listbox.size()):
            display_text = self.channel_listbox.get(i)
            is_completed = display_text.startswith(COMPLETED_PREFIX)
            url = display_text.replace(COMPLETED_PREFIX, "") if is_completed else display_text
            new_channels_config.append({"url": url, "completed": is_completed})
        self.config['channels'] = new_channels_config
        
        # --- MODIFICATION: Handle 'random' for delay field ---
        for key, var in self.config_vars.items():
            value = var.get().strip()
            if key == 'delay_between_requests' and value.lower() == 'random':
                self.config[key] = 'random'
            else:
                try:
                    self.config[key] = int(value)
                except ValueError:
                    try:
                        self.config[key] = float(value)
                    except ValueError:
                        self.config[key] = value
        
        save_config(self.config)
        if not silent:
            self.log_to_console(("SUCCESS", "💾 Configuration saved."))

    def add_channel(self):
        new_channel = simpledialog.askstring("Add Channel", "Enter new YouTube channel URL:")
        if new_channel and ('youtube.com/' in new_channel or 'youtu.be/' in new_channel):
            self.channel_listbox.insert(tk.END, new_channel)
            self.log_to_console(("INFO", f"➕ Added channel: {new_channel}"))
            self.save_config_from_ui()
        elif new_channel:
            messagebox.showwarning("Invalid URL", "Please enter a valid YouTube channel URL.")

    def remove_channel(self):
        selected_indices = self.channel_listbox.curselection()
        if not selected_indices: return
        for i in reversed(selected_indices):
            self.log_to_console(("INFO", f"➖ Removed channel: {self.channel_listbox.get(i)}"))
            self.channel_listbox.delete(i)
        self.save_config_from_ui()

    def move_item(self, direction):
        selected_indices = self.channel_listbox.curselection()
        if not selected_indices: return
        indices = selected_indices if direction == -1 else reversed(selected_indices)
        for i in indices:
            if 0 <= i + direction < self.channel_listbox.size():
                text = self.channel_listbox.get(i)
                color = self.channel_listbox.itemcget(i, "bg")
                self.channel_listbox.delete(i)
                self.channel_listbox.insert(i + direction, text)
                self.channel_listbox.itemconfig(i + direction, {'bg': color})
                self.channel_listbox.select_set(i + direction)
        if selected_indices: self.save_config_from_ui()

    def move_channel_up(self): self.move_item(-1)
    def move_channel_down(self): self.move_item(1)

    def log_to_console(self, msg_tuple): self.log_queue.put(msg_tuple)
    
    def mark_channel_as_completed(self, channel_url):
        all_items = self.channel_listbox.get(0, tk.END)
        for i, item in enumerate(all_items):
            clean_item = item.replace(COMPLETED_PREFIX, "")
            if clean_item == channel_url and not item.startswith(COMPLETED_PREFIX):
                self.channel_listbox.delete(i)
                self.channel_listbox.insert(i, f"{COMPLETED_PREFIX}{channel_url}")
                self.channel_listbox.itemconfig(i, {'bg': '#2E4034', 'fg': '#A9D1A9'})
                self.save_config_from_ui()
                break

    def process_log_queue(self):
        try:
            while True:
                msg_type, msg_content = self.log_queue.get_nowait()
                self.console.config(state='normal')
                if msg_type == "PROGRESS":
                    self.console.delete("end-2l", "end-1c")
                    self.console.insert(tk.END, msg_content + '\n', "INFO")
                elif msg_type == "MARK_COMPLETE":
                    self.mark_channel_as_completed(msg_content)
                else:
                    self.console.insert(tk.END, msg_content + '\n', msg_type)
                self.console.config(state='disabled')
                self.console.see(tk.END)
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self.process_log_queue)

    def update_button_states(self):
        is_running, is_paused = self.is_running.is_set(), self.is_paused.is_set()
        
        idle_state = tk.NORMAL if not is_running else tk.DISABLED
        running_state = tk.NORMAL if is_running and not is_paused else tk.DISABLED
        
        self.start_button.config(state=tk.NORMAL if not is_running or is_paused else tk.DISABLED)
        self.start_button.config(text="Resume" if is_paused else "Start Processing")
        self.pause_button.config(state=running_state)
        self.stop_button.config(state=tk.NORMAL if is_running else tk.DISABLED)
        self.save_button.config(state=idle_state)
        
        for btn in self.channel_buttons:
            btn.config(state=idle_state)
        for entry in self.config_entries:
            entry.config(state=idle_state)

    def start_processing(self):
        if self.is_paused.is_set():
            self.is_paused.clear()
            return
        if self.is_running.is_set(): return
        self.save_config_from_ui()
        completion_callback = lambda url: self.log_to_console(("MARK_COMPLETE", url))
        processor = TranscriptProcessor(self.config.copy(), self.log_to_console, self.is_paused, self.is_running, completion_callback)
        self.worker_thread = threading.Thread(target=processor.run, daemon=True)
        self.worker_thread.start()
        self.master.after(100, self.check_worker_status)

    def pause_processing(self):
        if self.is_running.is_set(): self.is_paused.set()

    def stop_processing(self):
        if not self.is_running.is_set(): return
        if messagebox.askokcancel("Stop Processing", "Are you sure you want to stop?\nThe process will halt after its current task and all progress will be saved."):
            self.log_to_console(("WARNING", "🛑 Stop requested by user. Please wait for current task to finish..."))
            self.is_running.clear()
            if self.is_paused.is_set():
                self.is_paused.clear() # Un-pause to allow thread to see stop signal
            self.stop_button.config(state=tk.DISABLED) # Prevent multiple clicks

    def check_worker_status(self):
        self.update_button_states()
        if self.worker_thread and self.worker_thread.is_alive():
            self.master.after(500, self.check_worker_status)
        else:
            if self.is_running.is_set():
                self.is_running.clear()
            self.update_button_states()

    def on_closing(self):
        if self.is_running.is_set():
            if messagebox.askokcancel("Quit", "A process is running. Stop and quit?"):
                self.is_running.clear()
                if self.worker_thread: self.worker_thread.join(timeout=1.5)
                self.master.destroy()
        else:
            self.save_config_from_ui()
            self.master.destroy()

if __name__ == '__main__':
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()