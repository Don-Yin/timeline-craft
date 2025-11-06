# timeline for powerpoint
this project is about adding a timeline (with smooth transitions) in microsoft powerpoint slides.

this is the PyPI package version of the project (`timeline-craft`).

# demo
![demo](https://github.com/don-yin/powerpoint-timeline/blob/cf5610f7db48a2f3e2fb747e2f197c5dbedd45e8/public/demo.gif)

# preview compare
|      before processing       |      after processing      |
| :--------------------------: | :------------------------: |
| ![before](public/before.png) | ![after](public/after.png) |

# use

```python
from pathlib import Path
from pptx import Presentation
from timeline import set_sidebar_timeline, move_elements_to_right, set_morph_transitions, Configurations


if __name__ == "__main__":
    path_ppt = Path("~/example.pptx")
    tags = ["intro"] * 15 + ["methods"] * 10 + ["results"] * 8

    example_ppt = Presentation(path_ppt)

    config = Configurations(
        sidebar_width=0.12,
        sidebar_item_height=0.10,
        sidebar_init_font_size=18,
        transition_duration=0.3,
        apply_morph_transition=True,
    )

    move_elements_to_right(example_ppt, config=config)
    set_sidebar_timeline(ppt=example_ppt, tags=tags, config=config)
    set_morph_transitions(example_ppt, config=config)

    example_ppt.save(path_ppt.parent / "timeline.pptx")
```