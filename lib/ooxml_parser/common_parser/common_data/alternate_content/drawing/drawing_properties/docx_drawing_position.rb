# Docx Drawing Position
module OoxmlParser
  class DocxDrawingPosition
    attr_accessor :relative_from, :offset, :align

    def self.parse(drawing_position_node)
      position = DocxDrawingPosition.new
      position.relative_from = case drawing_position_node.attribute('relativeFrom').value
                               when 'leftMargin'
                                 :left_margin
                               when 'rightMargin'
                                 :right_margin
                               when 'bottomMargin'
                                 :bottom_margin
                               when 'topMargin'
                                 :top_margin
                               when 'insideMargin'
                                 :inside_margin
                               when 'outsideMargin'
                                 :outside_margin
                               else
                                 drawing_position_node.attribute('relativeFrom').value.to_sym
                               end
      drawing_position_node.xpath('*').each do |position_node_child|
        case position_node_child.name
        when 'posOffset'
          position.offset = (position_node_child.text.to_f / 360_000.0).round(3)
        when 'align'
          position.align = position_node_child.text.to_sym
        end
      end
      position
    end
  end
end
