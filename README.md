# hoichoi Content Licensing Microsite

## Structure
```
/
├── build.py              ← Main build script (design lives here)
├── requirements.txt      ← Python dependencies
├── vercel.json           ← Vercel build config
├── data/                 ← Excel source files
│   ├── hoichoi_series_data_0513.xlsx
│   ├── hoichoi_ip_adaptation_data_0513.xlsx
│   └── hoichoi_movies_data_0513_2.xlsx
├── logos/                ← Brand logos (PNG, transparent bg)
│   ├── hoichoi.png
│   └── netflix.png
└── output/               ← Generated HTML (Vercel serves this)
```

## Workflow

### Data change (Excel update)
1. Edit the relevant Excel file in `data/`
2. `git add . && git commit -m "update data" && git push`
3. Vercel rebuilds automatically (~60 seconds)

### Design change
1. Get updated `build.py` from Claude
2. Replace `build.py` in this repo
3. `git add . && git commit -m "design update" && git push`
4. Vercel rebuilds with new design applied to all data

## Pages
| File | Description |
|------|-------------|
| `output/index.html` | Master landing — hoichoi × Netflix cover + 3 blocks |
| `output/series.html` | Hindi Dubbed Originals — 23 titles paginated |
| `output/adaptations.html` | IP Remakes & Production — 9 IPs paginated |
| `output/movies_landing.html` | Bengali Films landing — 3 movie cards |
| `output/hoichoi_movies.html` | Bengali Films catalog — 3 movies paginated |

## Local build
```bash
pip install -r requirements.txt
python build.py
```
Open `output/index.html` in browser to preview.
