require_relative 'text_box/text_box_content'

module OoxmlParser
  # Class for working with TextBox (v:textbox)
  class TextBox < OOXMLDocumentObject
    # @return [TextBoxContent] text box content
    attr_reader :text_box_content

    # Parse TextBox object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TextBox] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'txbxContent'
          @text_box_content = TextBoxContent.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
