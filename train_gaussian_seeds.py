"""Train the Gaussian handover reward with several algorithms and seeds.

distance_score = exp(-(distance / 0.15) ** 2)
The other terms match custom_reward in hri_handover_tutorial.ipynb.
"""

import argparse
import time
from pathlib import Path

import numpy as np
from stable_baselines3 import PPO, SAC, TD3
from stable_baselines3.common.monitor import Monitor

from train_distance_scores import HandoverGymEnv, SCORES, SuccessCallback, make_reward

SCORES_HERE = {
    "gaussian": SCORES["gaussian"],
    "squared": lambda distance: max(0.0, (0.35 - distance) / 0.35) ** 2,
}


def make_model(algo, env, seed):
    common = dict(
        policy="MlpPolicy",
        env=env,
        learning_rate=3e-4,
        seed=seed,
        device="cpu",
        verbose=0,
    )
    if algo == "sac":
        return SAC(
            **common,
            buffer_size=200_000,
            batch_size=256,
            train_freq=1,
            gradient_steps=1,
            learning_starts=1000,
            ent_coef=0.005,
        )
    if algo == "td3":
        return TD3(
            **common,
            buffer_size=200_000,
            batch_size=256,
            train_freq=1,
            gradient_steps=1,
        )
    if algo == "ppo":
        return PPO(**common, n_steps=2048, batch_size=64, n_epochs=10)
    raise ValueError(algo)


def evaluate(model, reward_fn, episodes):
    env = HandoverGymEnv(reward_fn)
    successes = 0
    failures = 0
    for ep in range(episodes):
        obs, _ = env.reset(seed=ep)
        done = False
        info = {}
        dists = []
        speeds = []
        while not done:
            action, _ = model.predict(obs, deterministic=True)
            obs, _, terminated, truncated, info = env.step(action)
            done = terminated or truncated
            dists.append(float(np.linalg.norm(info["handle_pos"] - info["ellipsoid_pos"])))
            speeds.append(float(info["hammer_speed"]))
        dists = np.asarray(dists)
        speeds = np.asarray(speeds)
        success = bool(info.get("success", False))
        failure = bool(info.get("failure", False))
        successes += int(success)
        failures += int(failure)
        print(
            f"eval {ep} success {int(success)} failure {int(failure)} "
            f"hold {info.get('hold_steps', 0)} "
            f"min_cm {dists.min() * 100:.1f} final_cm {dists[-1] * 100:.1f} "
            f"last30_cm {dists[-30:].mean() * 100:.1f} "
            f"speed_mm {speeds[-30:].mean() * 1000:.1f}",
            flush=True,
        )
    env.close()
    return successes, failures


def train_one(algo, seed, timesteps, out_dir, eval_episodes, score_name):
    out_dir.mkdir(parents=True, exist_ok=True)
    name = f"{algo}_s{seed}"
    log_path = out_dir / "episodes.txt"
    reward_fn = make_reward(SCORES_HERE[score_name])
    train_env = Monitor(HandoverGymEnv(reward_fn))
    sac_setup = " learning_starts 1000 ent_coef 0.005" if algo == "sac" else ""
    log_path.write_text(
        f"{name} {score_name} timesteps {timesteps}{sac_setup}\n"
    )
    model = make_model(algo, train_env, seed)
    callback = SuccessCallback(log_path, name)
    started = time.time()
    model.learn(total_timesteps=timesteps, callback=callback, progress_bar=False)
    model.save(str(out_dir / "model"))
    train_env.close()
    successes, failures = evaluate(model, reward_fn, eval_episodes)
    elapsed = time.time() - started
    summary = (
        f"{name} timesteps {model.num_timesteps} "
        f"eval_success {successes}/{eval_episodes} "
        f"eval_failure {failures}/{eval_episodes} seconds {elapsed:.0f}\n"
    )
    (out_dir / "summary.txt").write_text(summary)
    print(summary, end="", flush=True)
    return summary


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--algo", choices=("sac", "td3", "ppo"), required=True)
    parser.add_argument("--seed", type=int, required=True)
    parser.add_argument("--timesteps", type=int, default=120_000)
    parser.add_argument("--eval-episodes", type=int, default=5)
    parser.add_argument("--score", choices=tuple(SCORES_HERE), default="gaussian")
    parser.add_argument("--out", type=Path, default=None)
    args = parser.parse_args()
    out = args.out or Path(f"saved_models/{args.score}_seeds")
    train_one(
        args.algo,
        args.seed,
        args.timesteps,
        out / f"{args.algo}_s{args.seed}",
        args.eval_episodes,
        args.score,
    )


if __name__ == "__main__":
    main()
