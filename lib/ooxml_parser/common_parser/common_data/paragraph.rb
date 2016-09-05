require_relative 'paragraph/paragraph_properties'
require_relative 'paragraph/paragraph_run'
require_relative 'paragraph/text_field'
module OoxmlParser
  class Paragraph
    attr_accessor :properties, :runs, :text_field, :formulas

    def initialize(runs = [], formulas = [])
      @runs = runs
      @formulas = formulas
      @runs = []
    end

    alias characters runs
    alias character_style_array runs
    alias characters= runs=
    alias character_style_array= runs=

    def self.parse(paragraph_node)
      paragraph = Paragraph.new
      paragraph_node.xpath('*').each do |paragraph_node_child|
        case paragraph_node_child.name
        when 'pPr'
          paragraph.properties = ParagraphProperties.new(parent: paragraph).parse(paragraph_node_child)
        when 'fld'
          paragraph.text_field = TextField.parse(paragraph_node_child)
        when 'r'
          paragraph.characters << ParagraphRun.parse(paragraph_node_child)
        when 'm'
          # TODO: add parsing formulas in paragraph
        end
      end
      paragraph
    end
  end
end
