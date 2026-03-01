# Financial Researcher

A CrewAI-powered financial research agent that analyzes companies, generates comprehensive research reports in Markdown, and automatically converts them into professional PowerPoint presentations.

## Features

- 🔍 **Intelligent Research**: Uses AI agents to gather and analyze financial data
- 📊 **Comprehensive Reports**: Generates detailed markdown reports with key insights
- 🎯 **Professional Presentations**: Automatically converts reports to beautifully formatted PowerPoint slides
- 🤖 **Multi-Agent System**: Leverages CrewAI's crew framework for collaborative research

## Project Structure

```
financial_researcher/
├── src/
│   └── financial_researcher/
│       ├── __init__.py
│       ├── main.py              # Main entry point & crew orchestration
│       ├── crew.py              # Crew configuration & agent definitions
│       ├── tasks.py             # Task definitions for agents
│       ├── tools.py             # Custom tools & utilities
│       ├── ppt_generator.py     # PowerPoint generation from markdown
│       └── config/
│           ├── agents.yaml      # Agent configurations
│           └── tasks.yaml       # Task configurations
├── output/
│   ├── report.md               # Generated markdown report
│   └── report.pptx             # Generated PowerPoint presentation
├── .env                        # Environment variables (API keys)
├── pyproject.toml              # Project dependencies & metadata
├── README.md                   # This file
└── .gitignore                  # Git ignore rules
```

## Installation

### Prerequisites
- Python 3.10+
- pip or uv package manager

### Setup Steps

1. **Navigate to project directory:**
   ```bash
   cd financial_researcher
   ```

2. **Create & activate virtual environment:**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
   
   Or if using uv:
   ```bash
   uv pip install -r requirements.txt
   ```

4. **Set up environment variables:**
   ```bash
   cp .env.example .env
   # Edit .env and add your API keys (Anthropic, OpenAI, etc.)
   ```

## Usage

### Run the Financial Researcher

```bash
crewai run
```

Or directly with Python:

```bash
python src/financial_researcher/main.py
```

### Output

The crew will generate two files in the `output/` folder:

- **report.md** - Detailed markdown research report
- **report.pptx** - Professional PowerPoint presentation

## Configuration

### Agents Configuration (`config/agents.yaml`)
Define your research agents and their capabilities

### Tasks Configuration (`config/tasks.yaml`)
Define research tasks and expected outputs

### Environment Variables (`.env`)
```
ANTHROPIC_API_KEY=your_key_here
OPENAI_API_KEY=your_key_here
```

## Technologies Used

- **CrewAI** - Multi-agent orchestration framework
- **python-pptx** - PowerPoint generation
- **Anthropic Claude** - Language model
- **Python 3.10+** - Programming language

## Dependencies

Key packages:
- `crewai` - Agentic framework
- `python-pptx` - PowerPoint creation
- `python-dotenv` - Environment management
- Additional dependencies in `pyproject.toml`

## Contributing

1. Create a feature branch
2. Make your changes
3. Test thoroughly
4. Submit a pull request

## License

This project is licensed under the MIT License.

## Notes

- Ensure API keys are set in `.env` before running
- First run may take longer as models initialize
- Reports are saved in `output/` directory
- PowerPoint styling is automatically applied
