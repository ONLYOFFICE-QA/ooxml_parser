require_relative 'paragraph_run/run_properties'
require_relative 'paragraph_run/text'
module OoxmlParser
  # Class for parsing `r` tags
  class ParagraphRun < OOXMLDocumentObject
    attr_accessor :properties, :text
    # @return [Text] text of run
    attr_reader :t
    # @return [Tab] tab of paragraph
    attr_reader :tab

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
        when 'tab'
          @tab = Tab.new(parent: self).parse(node_child)
        end
      end
      self
    end

    def instruction
      parent.instruction
    end

    def page_number
      parent.page_numbering?
    end

    def link
      parent.parent.hyperlink
    end

    extend Gem::Deprecate
    deprecate :instruction, 'parent.instruction', 2020, 1
    deprecate :page_number, 'parent.page_numbering?', 2020, 1
    deprecate :link, 'parent.parent.hyperlink', 2020, 1
  end
end
