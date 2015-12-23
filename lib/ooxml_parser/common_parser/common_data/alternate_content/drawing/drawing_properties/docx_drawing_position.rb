# Docx Drawing Position
module OoxmlParser
  class DocxDrawingPosition
    attr_accessor :relative_from, :offset, :align

    def self.parse(drawing_position_node)
      position = DocxDrawingPosition.new
      case drawing_position_node.attribute('relativeFrom').value
      when 'leftMargin'
        position.relative_from = :left_margin
      when 'rightMargin'
        position.relative_from = :right_margin
      when 'bottomMargin'
        position.relative_from = :bottom_margin
      when 'topMargin'
        position.relative_from = :top_margin
      when 'insideMargin'
        position.relative_from = :inside_margin
      when 'outsideMargin'
        position.relative_from = :outside_margin
      else
        position.relative_from = drawing_position_node.attribute('relativeFrom').value.to_sym
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
