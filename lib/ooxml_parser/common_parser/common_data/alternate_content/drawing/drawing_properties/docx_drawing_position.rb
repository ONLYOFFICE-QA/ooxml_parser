module OoxmlParser
  # Docx Drawing Position
  class DocxDrawingPosition < OOXMLDocumentObject
    attr_accessor :relative_from, :offset, :align

    # Parse DocxDrawingPosition object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxDrawingPosition] result of parsing
    def parse(node)
      @relative_from = case node.attribute('relativeFrom').value
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
                         node.attribute('relativeFrom').value.to_sym
                       end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'posOffset', 'pctPosHOffset', 'pctPosVOffset'
          @offset = OoxmlSize.new(node_child.text.to_f, :emu)
        when 'align'
          @align = node_child.text.to_sym
        end
      end
      self
    end
  end
end
