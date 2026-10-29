"""Train one SAC policy for each notebook distance score.

The rest of the reward matches custom_reward in hri_handover_tutorial.ipynb.
Success is the environment flag: the handle stays inside 2 cm at <= 5 mm/s
for 60 steps.
"""

import argparse
import sys
import time
from pathlib import Path

import numpy as np

REPO = Path(__file__).resolve().parent / "mujoco" / "robosuite"
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

import gymnasium as gym
import robosuite as suite
from gymnasium import spaces
from stable_baselines3 import SAC
from stable_baselines3.common.callbacks import BaseCallback
from stable_baselines3.common.monitor import Monitor

SCORES = {
    "exponential": lambda distance: float(np.exp(-distance / 0.07)),
    "gaussian": lambda distance: float(np.exp(-(distance / 0.15) ** 2)),
    "inverse": lambda distance: 0.003 / (0.003 + distance),
    "lorentz": lambda distance: 1.0 / (1.0 + (distance / 0.03) ** 2),
}


def make_reward(score_fn):
    def reward_fn(obs, info):
        del obs
        distance = float(np.linalg.norm(info["handle_pos"] - info["ellipsoid_pos"]))
        speed = float(info["hammer_speed"])
        hold_steps = info["hold_steps"]
        distance_score = score_fn(distance)
        reward = distance_score
        if distance <= 0.02:
            reward = reward + 1.0
        reward = reward + distance_score * 0.003 / max(speed, 0.003)
        if distance <= 0.02 and speed <= 0.005:
            reward = reward + 1.0
            if hold_steps >= 60:
                reward = reward + 10.0
        return float(reward)

    return reward_fn


class HandoverGymEnv(gym.Env):
    def __init__(self, reward_fn, horizon=300):
        super().__init__()
        self._env = suite.make(
            env_name="HumanRobotHandover",
            robots="XArm7Handover",
            has_renderer=False,
            has_offscreen_renderer=False,
            use_camera_obs=False,
            reward_shaping=True,
            horizon=horizon,
            control_freq=20,
            initialization_noise=None,
            seed=0,
            hard_reset=False,
        )
        self._env.start_holding = True
        self._reward_fn = reward_fn
        self._low, self._high = self._env.action_spec
        dummy_obs = self._env.reset()
        self._obs_keys = sorted(k for k in dummy_obs.keys() if not k.endswith("_image"))
        flat_obs = self._flatten_obs(dummy_obs)
        self.observation_space = spaces.Box(
            low=-np.inf, high=np.inf, shape=flat_obs.shape, dtype=np.float32
        )
        self.action_space = spaces.Box(
            low=self._low.astype(np.float32),
            high=self._high.astype(np.float32),
            dtype=np.float32,
        )

    def _flatten_obs(self, obs):
        parts = [np.asarray(obs[k], dtype=np.float32).ravel() for k in self._obs_keys]
        return np.concatenate(parts)

    def reset(self, *, seed=None, options=None):
        if seed is not None:
            np.random.seed(seed)
        obs = self._env.reset()
        return self._flatten_obs(obs), {}

    def step(self, action):
        obs, env_reward, done, info = self._env.step(action)
        reward = self._reward_fn(obs, info)
        terminated = bool(info.get("success", False)) or bool(info.get("failure", False))
        truncated = bool(done) and not terminated
        return self._flatten_obs(obs), float(reward), terminated, truncated, info

    def close(self):
        self._env.close()


class SuccessCallback(BaseCallback):
    def __init__(self, log_path, name):
        super().__init__()
        self.log_path = Path(log_path)
        self.name = name
        self.episodes = 0
        self.successes = 0
        self.failures = 0

    def _on_step(self):
        if self.num_timesteps % 1000 == 0:
            print(
                f"{self.name} step {self.num_timesteps} "
                f"episodes {self.episodes} successes {self.successes}",
                flush=True,
            )
        for info in self.locals.get("infos", []):
            if not info.get("episode"):
                continue
            self.episodes += 1
            success = bool(info.get("success", False))
            failure = bool(info.get("failure", False))
            self.successes += int(success)
            self.failures += int(failure)
            line = (
                f"ep {self.episodes} success {int(success)} failure {int(failure)} "
                f"reward {info['episode']['r']:.3f} len {info['episode']['l']:.0f} "
                f"total_success {self.successes}\n"
            )
            with self.log_path.open("a") as f:
                f.write(line)
            print(line, end="", flush=True)
        return True


def evaluate(model, reward_fn, episodes):
    env = HandoverGymEnv(reward_fn)
    successes = 0
    failures = 0
    for ep in range(episodes):
        obs, _ = env.reset(seed=ep)
        done = False
        info = {}
        while not done:
            action, _ = model.predict(obs, deterministic=True)
            obs, _, terminated, truncated, info = env.step(action)
            done = terminated or truncated
        successes += int(bool(info.get("success", False)))
        failures += int(bool(info.get("failure", False)))
        print(
            f"eval {ep} success {int(bool(info.get('success', False)))} "
            f"failure {int(bool(info.get('failure', False)))} "
            f"hold {info.get('hold_steps', 0)}",
            flush=True,
        )
    env.close()
    return successes, failures


def train_one(name, timesteps, out_dir, eval_episodes, resume):
    out_dir.mkdir(parents=True, exist_ok=True)
    log_path = out_dir / "episodes.txt"
    reward_fn = make_reward(SCORES[name])
    train_env = Monitor(HandoverGymEnv(reward_fn))
    checkpoint = out_dir / "sac.zip"
    if resume:
        if not checkpoint.exists():
            raise SystemExit(f"missing checkpoint {checkpoint}")
        model = SAC.load(str(out_dir / "sac"), env=train_env, device="cpu")
        # The zip has the weights, not the replay buffer. Keep acting with the
        # loaded policy, and skip updates until a new batch has been collected.
        original_train = model.train

        def train_after_batch(*args, **kwargs):
            if model.replay_buffer.size() < model.batch_size:
                return None
            return original_train(*args, **kwargs)

        model.train = train_after_batch
        prior = sum(1 for line in log_path.read_text().splitlines() if line.startswith("ep "))
        with log_path.open("a") as f:
            f.write(f"resume at step {model.num_timesteps} toward {timesteps}\n")
        print(f"{name} resume at step {model.num_timesteps}", flush=True)
    else:
        log_path.write_text("")
        prior = 0
        model = SAC(
            policy="MlpPolicy",
            env=train_env,
            learning_rate=3e-4,
            buffer_size=200_000,
            batch_size=256,
            train_freq=1,
            gradient_steps=1,
            learning_starts=1000,
            ent_coef=0.005,
            verbose=0,
            seed=42,
            device="cpu",
        )
    callback = SuccessCallback(log_path, name)
    callback.episodes = prior
    started = time.time()
    model.learn(
        total_timesteps=timesteps,
        callback=callback,
        progress_bar=False,
        reset_num_timesteps=not resume,
    )
    model.save(str(out_dir / "sac"))
    model.save_replay_buffer(str(out_dir / "replay"))
    train_env.close()
    successes, failures = evaluate(model, reward_fn, eval_episodes)
    elapsed = time.time() - started
    summary = (
        f"{name} timesteps {model.num_timesteps} eval_success {successes}/{eval_episodes} "
        f"eval_failure {failures}/{eval_episodes} seconds {elapsed:.0f}\n"
    )
    with (out_dir / "summary.txt").open("a") as f:
        f.write(summary)
    print(summary, end="", flush=True)
    return summary


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--name", choices=list(SCORES), default=None)
    parser.add_argument("--timesteps", type=int, default=20_000)
    parser.add_argument("--eval-episodes", type=int, default=5)
    parser.add_argument("--out", type=Path, default=Path("saved_models/distance_scores"))
    parser.add_argument("--resume", action="store_true")
    args = parser.parse_args()
    names = [args.name] if args.name else list(SCORES)
    for name in names:
        train_one(name, args.timesteps, args.out / name, args.eval_episodes, args.resume)


if __name__ == "__main__":
    main()
