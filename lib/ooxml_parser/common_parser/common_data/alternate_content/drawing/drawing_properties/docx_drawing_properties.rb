require_relative 'ooxml_coordinates'
require_relative 'ooxml_size'
require_relative 'docx_drawing_distance_from_text'
require_relative 'docx_drawing_position'
require_relative 'docx_wrap_drawing'
# Docx Drawing Properties
module OoxmlParser
  class DocxDrawingProperties
    attr_accessor :distance_from_text, :simple_position, :wrap, :object_size, :vertical_position, :horizontal_position,
                  :relative_height
    attr_accessor :size_relative_horizontal
    attr_accessor :size_relative_vertical
  end
end
