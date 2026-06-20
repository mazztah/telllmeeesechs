FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    git ffmpeg libsm6 libxext6 gcc g++ \
    portaudio19-dev libsndfile1 libassimp-dev \
    && rm -rf /var/lib/apt/lists/*

RUN useradd -m -u 1000 user

# WICHTIG: Verzeichnis VOR dem USER-Wechsel erstellen und chownen!
RUN mkdir -p /home/user/app && chown user:user /home/user/app

USER user
WORKDIR /home/user/app

ENV PATH="/home/user/.local/bin:$PATH"
ENV DATA_DIR="/data"

# WICHTIG: Port dynamisch von Cloud Run übernehmen
ENV PORT=8080

COPY --chown=user:user requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

COPY --chown=user:user . .

EXPOSE 8080

# Dynamischer Port!
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT} --workers 1 --log-level info"]
