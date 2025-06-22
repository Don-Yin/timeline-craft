
# Timeline for PowerPoint

This project is about adding a timeline (with smooth transitions) in Microsoft PowerPoint slides.

## Demo
![Demo](https://github.com/Don-Yin/powerpoint-timeline/blob/cf5610f7db48a2f3e2fb747e2f197c5dbedd45e8/public/demo.gif)

## Preview Compare
|      Before Processing       |      After Processing      |
| :--------------------------: | :------------------------: |
| ![Before](public/before.png) | ![After](public/after.png) |


## Use

```python
from pathlib import Path
from pptx import Presentation
from timeline import set_sidebar_timeline, move_elements_to_right


if __name__ == "__main__":
    path_ppt = Path("/Users/donyin/Desktop/example.pptx")
    tags = ["Intro"] * 33

    example_ppt = Presentation(path_ppt)

    move_elements_to_right(example_ppt, sidebar_width=0.12)
    set_sidebar_timeline(ppt=example_ppt, tags=tags)

    example_ppt.save(path_ppt.parent / "timeline.pptx")
```

# Disclaimer
This project is under development, with further refinements and feature additions in progress. Your feedback and contributions are highly appreciated. This is probably the dirtiest piece of code I have written, and I blame it on python-pptx, on which this project, out of no other choice, is based. Although useful it is a horrible package with irregular syntax and poor documentation.
