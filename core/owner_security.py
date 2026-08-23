import os
import json
import time
import secrets
import logging
from collections import defaultdict
from werkzeug.security import generate_password_hash, check_password_hash

logger = logging.getLogger('sebn-maintenance')

class RateLimiter:
    """In-memory sliding window rate limiter with lockout support."""
    def __init__(self):
        self.attempts = defaultdict(list)
        self.lockouts = {}

    def is_rate_limited(self, key, action, max_attempts=5, window_seconds=900, lockout_seconds=900):
        now = time.time()
        lock_key = f"{key}:{action}"

        # Check if in active lockout
        if lock_key in self.lockouts:
            unlock_time = self.lockouts[lock_key]
            if now < unlock_time:
                remaining_seconds = int(unlock_time - now)
                return True, remaining_seconds
            else:
                del self.lockouts[lock_key]
                self.attempts[lock_key] = []

        # Filter attempts within the window
        history = [t for t in self.attempts[lock_key] if now - t < window_seconds]
        self.attempts[lock_key] = history

        if len(history) >= max_attempts:
            self.lockouts[lock_key] = now + lockout_seconds
            return True, lockout_seconds

        return False, 0

    def record_attempt(self, key, action):
        now = time.time()
        lock_key = f"{key}:{action}"
        self.attempts[lock_key].append(now)

    def reset(self, key, action):
        lock_key = f"{key}:{action}"
        self.attempts.pop(lock_key, None)
        self.lockouts.pop(lock_key, None)


class OwnerSecurityManager:
    """
    Dedicated security manager for Owner-Only Administrator Recovery.
    Completely isolated from regular user and administrator accounts.
    """
    def __init__(self, data_dir: str):
        self.data_dir = data_dir
        self.config_path = os.path.join(data_dir, "owner_config.json")
        self.audit_log_path = os.path.join(data_dir, "security_audit.log")
        self.rate_limiter = RateLimiter()
        self.active_reset_tokens = {}  # token -> expiration_timestamp
        
        self._ensure_config()

    def _ensure_config(self):
        """Initializes owner configuration with secure password hashing if not present."""
        if not os.path.exists(self.config_path):
            default_username = os.environ.get("SEBN_OWNER_USERNAME", "owner")
            default_password = os.environ.get("SEBN_OWNER_PASSWORD", "Owner@SEBN2026!")
            
            # Prefer env hash if provided, otherwise generate secure hash
            env_hash = os.environ.get("SEBN_OWNER_PASSWORD_HASH")
            password_hash = env_hash if env_hash else generate_password_hash(default_password)

            data = {
                "owner_username": default_username,
                "owner_password_hash": password_hash,
                "admin_session_epoch": time.time(),
                "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "last_recovery": None
            }
            self._save_config(data)

    def _load_config(self) -> dict:
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"Error loading owner config: {e}")
        return {
            "owner_username": os.environ.get("SEBN_OWNER_USERNAME", "owner"),
            "owner_password_hash": generate_password_hash(os.environ.get("SEBN_OWNER_PASSWORD", "Owner@SEBN2026!")),
            "admin_session_epoch": time.time()
        }

    def _save_config(self, data: dict):
        os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

    def get_owner_username(self) -> str:
        # Check env first for container/cloud deployment override
        env_user = os.environ.get("SEBN_OWNER_USERNAME")
        if env_user:
            return env_user.strip()
        config = self._load_config()
        return config.get("owner_username", "owner")

    def get_owner_password_hash(self) -> str:
        env_hash = os.environ.get("SEBN_OWNER_PASSWORD_HASH")
        if env_hash:
            return env_hash
        env_pass = os.environ.get("SEBN_OWNER_PASSWORD")
        if env_pass:
            return generate_password_hash(env_pass)
        config = self._load_config()
        return config.get("owner_password_hash", "")

    def verify_owner_login(self, username: str, password: str) -> bool:
        """
        Verifies owner username and password.
        Uses secure constant-time password hash verification.
        """
        if not username or not password:
            return False

        stored_user = self.get_owner_username()
        stored_hash = self.get_owner_password_hash()

        if username.strip() != stored_user:
            return False

        if not stored_hash:
            return False

        return check_password_hash(stored_hash, password)

    def generate_reset_token(self, ttl_seconds: int = 600) -> str:
        """Generates a high-entropy single-use recovery token valid for ttl_seconds (default 10 min)."""
        now = time.time()
        # Clean expired tokens
        self.active_reset_tokens = {t: exp for t, exp in self.active_reset_tokens.items() if exp > now}
        
        token = secrets.token_urlsafe(32)
        self.active_reset_tokens[token] = now + ttl_seconds
        return token

    def is_reset_token_valid(self, token: str) -> bool:
        if not token:
            return False
        now = time.time()
        exp = self.active_reset_tokens.get(token)
        return exp is not None and exp > now

    def consume_reset_token(self, token: str):
        self.active_reset_tokens.pop(token, None)

    def get_admin_session_epoch(self) -> float:
        config = self._load_config()
        return float(config.get("admin_session_epoch", 0.0))

    def invalidate_all_admin_sessions(self):
        """Updates admin session epoch to now, invalidating all existing admin sessions across all browsers."""
        now = time.time()
        config = self._load_config()
        config["admin_session_epoch"] = now
        config["last_recovery"] = time.strftime("%Y-%m-%d %H:%M:%S")
        self._save_config(config)
        logger.info(f"AUTH SECURITY: All existing Administrator sessions invalidated at epoch {now}")

    def is_admin_session_valid(self, session_login_time: float) -> bool:
        """Checks if an admin session was authenticated after the last invalidation epoch."""
        if not session_login_time:
            return False
        return float(session_login_time) >= self.get_admin_session_epoch()

    def log_security_event(self, event_type: str, actor: str, ip_address: str, details: str):
        """Appends an immutable audit log entry for security and recovery events."""
        entry = {
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "event": event_type,
            "actor": actor,
            "ip": ip_address,
            "details": details
        }
        try:
            with open(self.audit_log_path, 'a', encoding='utf-8') as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception as e:
            logger.error(f"Failed to write security audit log: {e}")
