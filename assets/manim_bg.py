from manim import *  # type: ignore
import numpy as np

# Configurações para banner (recorte da área do print)
config.pixel_width = 1280
config.pixel_height = 360
config.frame_rate = 30
config.background_color = "#0b1117"  # tom escuro do app


class ContinuousMotion(Scene):
    def construct(self) -> None:
        def func(pos: np.ndarray) -> np.ndarray:
            return np.sin(pos[0] / 2.0) * UR + np.cos(pos[1] / 2.0) * LEFT

        stream_lines = StreamLines(
            func,
            stroke_width=2,
            max_anchors_per_line=30,
        )
        self.add(stream_lines)
        stream_lines.start_animation(warm_up=False, flow_speed=1.5)
        # Duração exata para que feche o ciclo e possamos loopar no front
        self.wait(stream_lines.virtual_time / stream_lines.flow_speed)


