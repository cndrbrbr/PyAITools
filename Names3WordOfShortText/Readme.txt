Get ollama 
https://hub.docker.com/r/ollama/ollama
docker run -d -v ollama:/root/.ollama -p 11434:11434 --name ollama ollama/ollama
Load LangModel
docker exec -it ollama ollama pull llama3.1:8b
http://localhost:11434/api/tags
