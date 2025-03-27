import vosk
import sounddevice as sd
import numpy as np
import json
import queue
import threading
import keyboard
import logging
import time
import difflib

# MediaPipe imports
import cv2
import mediapipe as mp
import numpy as np

import win32com.client
import win32gui
import win32api
import win32con

class WindowsGameInput:
    @staticmethod
    def activate_game_window(window_title):
        """Bring game window to foreground"""
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            # Bring window to foreground
            win32gui.SetForegroundWindow(hwnd)

    @staticmethod
    def send_key(key):
        """Send low-level key press"""
        # Map your keys to virtual key codes
        key_map = {
            'a': win32con.VK_LEFT,
            'd': win32con.VK_RIGHT,
            'w': win32con.VK_UP,
            's': win32con.VK_DOWN,
            'space': win32con.VK_SPACE,
            'enter': win32con.VK_RETURN,
            'tab': win32con.VK_TAB,
            'esc': win32con.VK_ESCAPE,
            'x': ord('X'),
            'e': ord('I'),
            'shift': win32con.VK_SHIFT
        }
        
        virtual_key = key_map.get(key)
        if virtual_key:
            # Simulate key press and release
            win32api.keybd_event(virtual_key, 0, 0, 0)  # Key down
            time.sleep(0.05)  # Short delay
            win32api.keybd_event(virtual_key, 0, win32con.KEYEVENTF_KEYUP, 0)  # Key up

class HeadMovementController:
    def __init__(self, 
                 tilt_threshold: float = 10.0, 
                 smoothing_factor: float = 0.7):
        """
        Initialize head movement tracking
        
        :param tilt_threshold: Degrees of tilt to trigger key press
        :param smoothing_factor: Smooth out rapid fluctuations
        """
        # MediaPipe setup
        self.mp_face_mesh = mp.solutions.face_mesh
        self.mp_drawing = mp.solutions.drawing_utils
        
        # Tracking parameters
        self.tilt_threshold = tilt_threshold
        self.smoothing_factor = smoothing_factor
        
        # Tracking control
        self.is_tracking = False
        self.tracking_thread = None
        
        # Logging
        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        
        # Smoothed head pose
        self.smoothed_roll = 0
        
        # Tilt event tracking
        self.last_tilt_direction = None
        self.tilt_event_count = 0
        self.max_consecutive_events = 5

    def calculate_head_tilt(self, landmarks):
        """
        Calculate head tilt angle from face mesh landmarks
        
        :param landmarks: MediaPipe face mesh landmarks
        :return: Roll angle (tilt)
        """
        # Select specific landmarks for head orientation
        left_ear = landmarks.landmark[454]
        right_ear = landmarks.landmark[234]
        
        # Calculate angle between ears
        dx = right_ear.x - left_ear.x
        dy = right_ear.y - left_ear.y
        
        # Calculate roll angle in degrees
        roll_angle = np.degrees(np.arctan2(dy, dx))
        if(roll_angle>0):
            roll_angle-=180
        elif roll_angle <0:
            roll_angle+=180
        return roll_angle

    def smooth_angle(self, current_angle):
        """
        Apply exponential smoothing to angle
        
        :param current_angle: Current head tilt angle
        :return: Smoothed angle
        """
        self.smoothed_roll = (
            self.smoothing_factor * current_angle + 
            (1 - self.smoothing_factor) * self.smoothed_roll
        )
        return self.smoothed_roll

    def track_head_movement(self):
        """
        Continuously track head movement and trigger key presses
        """
        # Open webcam
        cap = cv2.VideoCapture(0)
        
        # Ensure camera is opened
        if not cap.isOpened():
            self.logger.error("Could not open camera")
            return
        
        # MediaPipe Face Mesh setup
        with mp.solutions.face_mesh.FaceMesh(
            min_detection_confidence=0.5,
            min_tracking_confidence=0.5,
            max_num_faces=1
        ) as face_mesh:
            
            while self.is_tracking:
                # Read frame from webcam
                ret, frame = cap.read()
                if not ret:
                    self.logger.error("Failed to grab frame")
                    break
                
                # Flip frame for mirror effect
                frame = cv2.flip(frame, 1)
                
                # Convert BGR to RGB
                rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                
                # Process frame
                results = face_mesh.process(rgb_frame)
                
                if results.multi_face_landmarks:
                    for face_landmarks in results.multi_face_landmarks:
                        # Calculate head tilt
                        tilt_angle = self.calculate_head_tilt(face_landmarks)
                        
                        # Smooth the angle
                        smoothed_angle = self.smooth_angle(tilt_angle)
                        
                        # Determine tilt direction and key press
                        current_tilt_direction = None

                        if smoothed_angle > self.tilt_threshold:
                            current_tilt_direction = 'right'
                            try:
                                keyboard.press_and_release('d')
                                WindowsGameInput.send_key('d')
                            except Exception as e:
                                self.logger.error(f"Error pressing 'd': {e}")
                        elif smoothed_angle < -self.tilt_threshold:
                            current_tilt_direction = 'left'
                            try:
                                keyboard.press_and_release('a')
                                WindowsGameInput.send_key('a')
                            except Exception as e:
                                self.logger.error(f"Error pressing 'a': {e}")
                        
                        # Track consecutive tilt events
                        if current_tilt_direction:
                            if current_tilt_direction == self.last_tilt_direction:
                                self.tilt_event_count += 1
                            else:
                                self.tilt_event_count = 1
                            
                            # Log warning for excessive tilting
                            if self.tilt_event_count > self.max_consecutive_events:
                                self.logger.warning(f"Excessive {current_tilt_direction} tilting detected!")
                            
                            self.last_tilt_direction = current_tilt_direction
                
                # Draw tilt angle overlay
                cv2.putText(frame, 
                    f"Tilt: {self.smoothed_roll:.2f} degrees", 
                    (10, 30), 
                    cv2.FONT_HERSHEY_SIMPLEX, 
                    0.7, 
                    (0, 255, 0), 
                    2
                )
                
                # Show camera preview with larger window
                cv2.namedWindow('Head Movement Tracking', cv2.WINDOW_NORMAL)
                cv2.resizeWindow('Head Movement Tracking', 800, 600)
                cv2.imshow('Head Movement Tracking', frame)
                
                # Break loop if 'q' is pressed
                key = cv2.waitKey(1) & 0xFF
                if key == ord('q'):
                    break
        
        # Cleanup
        cap.release()
        cv2.destroyAllWindows()

    def start_tracking(self):
        """
        Start head movement tracking in a separate thread
        """
        if not self.is_tracking:
            self.is_tracking = True
            # Use daemon=True to ensure thread stops when main program exits
            self.tracking_thread = threading.Thread(
                target=self.track_head_movement,
                daemon=True
            )
            self.tracking_thread.start()
            self.logger.info("Head movement tracking started")

    def stop_tracking(self):
        """
        Stop head movement tracking
        """
        self.is_tracking = False
        if self.tracking_thread:
            self.tracking_thread.join(timeout=2)
            self.logger.info("Head movement tracking stopped")

class SimplifiedVoskSpeechController:
    def __init__(self, 
                 model_path: str,
                 word_options: list = None,
                 word_key_map: dict = None,
                 sample_rate: int = 16000,
                 similarity_threshold: float = 0.6):
        """
        Initialize simplified Vosk-based speech recognition controller
        
        :param model_path: Path to Vosk speech recognition model
        :param word_options: List of valid command words
        :param word_key_map: Mapping of words to keyboard actions
        :param sample_rate: Audio sample rate
        :param similarity_threshold: Threshold for word matching
        """
        # Setup logging
        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        
        # Vosk model setup
        try:
            self.model = vosk.Model("Vosk")
            self.recognizer = vosk.KaldiRecognizer(self.model, sample_rate)
        except Exception as e:
            self.logger.error(f"Failed to load Vosk model: {e}")
            raise
        
        # Word options and mapping
        self.word_options = word_options or [
            'up', 'down', 'left', 'right', 
            'jump', 'move', 'attack', 'punch', 
            'switch', 'enter', 'tab', 'escape','fuck'
        ]
        
        self.word_key_map = word_key_map or {
            'up': 'up',
            'down': 'down',
            'left': 'left',
            'right': 'right',
            'jump': 'space',
            'move': 'w',
            'attack': 'x',
            'punch': 'e',
            'switch': 'shift',
            'enter': 'enter',
            'tab': 'tab',
            'escape': 'esc',
            'start' : 'w',
        }
        
        # Similarity threshold for word matching
        self.similarity_threshold = similarity_threshold
        
        # Listening control
        self.is_listening = False
        self.listener_thread = None
        
        # Global flags
        self.accel = False

    def find_closest_word(self, spoken_text: str) -> str:
        """
        Find the closest matching word from predefined options
        
        :param spoken_text: Text recognized by Vosk
        :return: Closest matching word or None
        """
        # Split the text into words and match each
        for word in spoken_text.lower().split():
            # Use difflib to find closest match
            matches = difflib.get_close_matches(
                word, 
                self.word_options, 
                n=1, 
                cutoff=self.similarity_threshold
            )
            
            # Return first match if found
            if matches:
                return matches[0]
        
        return None

    def _process_recognized_text(self, text: str):
        """
        Process recognized text and trigger corresponding action
        
        :param text: Text recognized by Vosk
        """
        # Find closest matching word
        matched_word = self.find_closest_word(text)
        
        if matched_word:
            try:
                # Special handling for acceleration
                if matched_word == 'start':
                    self.accel = not self.accel

                else:
                    key = self.word_key_map.get(matched_word)
                    if key:
                        keyboard.press_and_release(key)
                        WindowsGameInput.send_key(key)
                        self.logger.info(f"Pressed key for word: {matched_word}")
            except Exception as e:
                self.logger.error(f"Error processing word {matched_word}: {e}")
        else:
            self.logger.warning(f"No matching word found for: {text}")

    def listen_and_control(self):
        """
        Continuously listen for words and control keyboard
        """
        def audio_callback(indata, frames, time, status):
            if status:
                self.logger.warning(status)
            
            if self.recognizer.AcceptWaveform(indata.tobytes()):
                result = json.loads(self.recognizer.Result())
                if 'text' in result and result['text'].strip():
                    self._process_recognized_text(result['text'])

        while self.is_listening:
            try:
                # Start audio stream
                with sd.InputStream(
                    samplerate=16000, 
                    channels=1, 
                    dtype='int16',
                    callback=audio_callback
                ):
                    print("Listening for voice commands...")
                    
                    # Keep the thread running
                    while self.is_listening:
                        # Optional: Check for acceleration
                        if self.accel:
                            try:
                                keyboard.press_and_release('w')
                                WindowsGameInput.send_key('w')
                            except Exception as e:
                                self.logger.error(f"Acceleration error: {e}")
                        
                        time.sleep(0.1)
            
            except Exception as e:
                self.logger.error(f"Unexpected error in audio stream: {e}")
                time.sleep(0.5)

    def start_listening(self):
        """
        Start listening for voice commands
        """
        if not self.is_listening:
            self.is_listening = True
            self.listener_thread = threading.Thread(
                target=self.listen_and_control,
                daemon=True
            )
            self.listener_thread.start()
            self.logger.info("Voice-controlled keyboard started")

    def stop_listening(self):
        """
        Stop listening for voice commands
        """
        self.is_listening = False
        if self.listener_thread:
            self.listener_thread.join(timeout=2)
            self.logger.info("Voice-controlled keyboard stopped")

    def add_word_option(self, word: str):
        """
        Add a new word option to the recognition list
        
        :param word: New word to add to options
        """
        if word.lower() not in self.word_options:
            self.word_options.append(word.lower())
            self.logger.info(f"Added new word option: {word}")

def main():
    # Replace with the actual path to your Vosk model
    model_path = "Vosk"
    
    # Create head movement controller
    head_controller = HeadMovementController(
        tilt_threshold=30, # Adjust sensitivity as needed
        smoothing_factor=0.7  # Smooth out rapid movements
    )
    
    # Create speech-to-keyboard controller
    speech_controller = SimplifiedVoskSpeechController(
        model_path=model_path,
        word_options=[
            'up', 'down', 'left', 'right', 
            'jump', 'move', 'attack', 'punch','start','fuck'
        ],
        similarity_threshold=0  # Adjust for looser or tighter matching
    )
    
    try:
        # Start head movement tracking
        head_controller.start_tracking()
        
        # Start speech recognition
        speech_controller.start_listening()
        
        # Keep main thread running
        while True:
            time.sleep(1)
    
    except KeyboardInterrupt:
        # Stop both controllers
        head_controller.stop_tracking()
        speech_controller.stop_listening()
        print("\nControllers stopped.")

if __name__ == "__main__":
    main()