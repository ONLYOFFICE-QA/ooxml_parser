require_relative 'paragraph_run/run_properties'
require_relative 'paragraph_run/text'
module OoxmlParser
  # Class for parsing `r` tags
  class ParagraphRun < OOXMLDocumentObject
    attr_accessor :properties, :text
    # @return [Text] text of run
    attr_reader :t

    def initialize(properties = RunProperties.new, text = '', parent: nil)
      @properties = properties
      @text = text
      @parent = parent
    end

    # Parse ParagraphRun object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ParagraphRun] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'rPr'
          @properties = RunProperties.new(parent: self).parse(node_child)
        when 't'
          @t = Text.new(parent: self).parse(node_child)
          @text = t.text
        end
      end
      self
    end
  end
end
