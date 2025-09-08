# PPT-automation-motivation.
Python script to auto-categorize motivational images and generate neatly formatted PowerPoint presentations.

## 📌 Project Overview
This project automates the process of **organizing motivational images** into predefined categories and then generates **PowerPoint presentations** for each category.  
Instead of manually sorting and creating slides, this script leverages AI (CLIP model)and Python automation to handle everything in seconds.


## ⚡ Features
- ✅ Automatically categorizes images into motivational themes:
  - Success & Hard Work  
  - Habits & Discipline  
  - Time & Productivity  
  - Attitude & Positivity  
  - Knowledge & Learning  
  - Life Lessons / General Motivation  
- ✅ Moves images into category-wise folders.  
- ✅ Creates **category-specific PPT files** with:
  - A title slide for each category  
  - All images centered and properly resized on their own slides  
- ✅ Works locally — no paid API needed.  

---

## 🛠️ Tech Stack
- Python 3.12+
- [PyTorch](https://pytorch.org/) + [OpenAI CLIP](https://github.com/openai/CLIP) → for image categorization  
- [Pillow](https://pillow.readthedocs.io/) → for image processing  
- [python-pptx](https://python-pptx.readthedocs.io/) → for creating PowerPoint slides  

---

## 🚀 How to Run

1. Clone this repository:
   ```bash
   git clone https://github.com/your-username/motivational-ppt-automation.git
   cd motivational-ppt-automation
2.Create a virtual environment (recommended):

  bash
  Copy code
  python -m venv .venv
  source .venv/bin/activate   # Mac/Linux
  .venv\Scripts\activate      # Windows
3.Install dependencies:
  
  bash
  Copy code
  pip install -r requirements.txt
4.Place your raw images inside the motivational_images/ folder.

5.Run the script:

  bash
  Copy code
  python start.py
  
6.Check the results:

  Categorized images → inside Final_images/
  
  Generated PPTs → inside presentations/

🌟 Future Improvements

🔍 Improve categorization with fine-tuned models.

☁️ Add option to use APIs (e.g., Gemini/OpenAI) for better text/image understanding.

📈 Generate a single master presentation with all categories.

🖼️ Support PDFs alongside PPTs.

🙌 Acknowledgements

OpenAI CLIP
 for image-text embeddings

python-pptx
 for PowerPoint automation
