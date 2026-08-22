from __future__ import annotations

from pathlib import Path

from office_shape_canvas import ShapeCanvas


class PictureProxy:
    def __init__(self, canvas: ShapeCanvas, shape) -> None:
        self.canvas = canvas
        self.shape = shape

    @property
    def name(self) -> str:
        return self.shape.name

    @name.setter
    def name(self, value: str) -> None:
        old_name = self.shape.name
        self.shape.name = value
        self.canvas._shape_index.pop(old_name, None)
        self.canvas._shape_index[value] = self.shape
        for op in reversed(self.canvas.ops):
            if op.values.get("name") == old_name:
                op.values["name"] = value
                break


class ShapesProxy:
    def __init__(self, canvas: ShapeCanvas) -> None:
        self.canvas = canvas

    @staticmethod
    def _points(value) -> float:
        return float(value.pt if hasattr(value, "pt") else value / 12700)

    def add_picture(self, path, left, top, width, height):
        shape = self.canvas.picture(
            Path(path),
            self._points(left),
            self._points(top),
            self._points(width),
            self._points(height),
        )
        return PictureProxy(self.canvas, shape)


class SlideProxy:
    def __init__(self, canvas: ShapeCanvas) -> None:
        self.canvas = canvas
        self.shapes = ShapesProxy(canvas)


class SlidesProxy:
    def __init__(self, slide: SlideProxy) -> None:
        self.slide = slide

    def add_slide(self, _layout):
        return self.slide
