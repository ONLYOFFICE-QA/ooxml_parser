# Docx Pattern Fill Data
module OoxmlParser
  class DocxPatternFill < OOXMLDocumentObject
    attr_accessor :foreground_color, :background_color, :preset

    def self.parse(fill_node)
      pattern = DocxPatternFill.new
      pattern.preset = fill_node.attribute('prst').value.to_sym
      fill_node.xpath('*').each do |fill_node_child|
        case fill_node_child.name
        when 'fgClr'
          pattern.foreground_color = Color.parse_color_model(fill_node_child)
        when 'bgClr'
          pattern.background_color = Color.parse_color_model(fill_node_child)
        end
      end
      pattern
    end
  end
end
