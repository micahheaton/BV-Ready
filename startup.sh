# startup.sh

# Install Python dependencies
pip install --no-cache-dir -r requirements.txt

# Run the Streamlit app
streamlit run app.py --server.port=$PORT