# Reddit to .docx

A simple python script that converts a reddit thread into a .docx document retaining the tree structure and some basic formatting. 

## Usage

`python reddit_to_docx.py -i "https://www.reddit.com/r/subreddit/comments/thread-id" -o "output_path/file_name.docx"`

## Requirements

- Python 3.3+
- `beautifulsoup4` library
- `lxml` library
- `python-docx` library
- `requests` library

## Notes

- Unsupported html tags are ignored
- Some subreddits use custom formatting via <a> tags, they are treated as hyperlinks just like the normal ones
- Tables are not yet supported and appear in a mangled form
- May take a couple of minutes to process threads with several thousand messages