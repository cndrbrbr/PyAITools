******************************************
* Convert short Text to 3 Word Summary
* Windows client with python and VS Code
* Docker Container running ollama
******************************************
Get ollama 
https://hub.docker.com/r/ollama/ollama
docker run -d -v ollama:/root/.ollama -p 11434:11434 --name ollama ollama/ollama
Load LangModel
docker exec -it ollama ollama pull llama3.1:8b
http://localhost:11434/api/tags
>>{"models":[{"name":"llama3.1:8b","model":"llama3.1:8b","modified_at":"2026-01-14T17:50:07.052855939Z","size":4920753328,"digest":"46e0c10c039e019119339687c3c1757cc81b9da49709a3b3924863ba87ca666e","details":{"parent_model":"","format":"gguf","family":"llama","families":["llama"],"parameter_size":"8.0B","quantization_level":"Q4_K_M"}}]}
Model Loaded
python.exe -m pip install --upgrade pip
python.exe -m pip install requests

┌──────────────────────────┐
│ Windows Computer         │
│                          │
│  Python Script           │
│   → http://localhost:11434
│                          │
└───────────▲──────────────┘
            │
┌───────────┴──────────────┐
│ Docker Container "ollama"│
│                          │
│  Ollama API Server       │
│  Model in /root/.ollama  │
│                          │
└──────────────────────────┘
