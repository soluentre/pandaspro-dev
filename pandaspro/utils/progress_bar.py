import sys
import time


def show_wait_progress_segment(segment_length, total_segments, current_segment, step=0.1):
    steps = int(segment_length / step)
    for i in range(steps + 1):
        progress = i / steps
        bar_length = 30
        block = int(round(bar_length * progress))

        overall_progress = ((current_segment - 1) / total_segments + progress / total_segments) * 100
        text = "\rProgress ({}/{}): [{}] {:.2f}%".format(current_segment, total_segments,
                                                         "#" * block + "-" * (bar_length - block), overall_progress)
        sys.stdout.write(text)
        sys.stdout.flush()
        time.sleep(step)
    sys.stdout.write("\n")


def show_wait_progress(total_wait_time, segments=5):
    segment_length = total_wait_time / segments
    for segment in range(1, segments + 1):
        show_wait_progress_segment(segment_length, segments, segment)