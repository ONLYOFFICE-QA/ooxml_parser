require_relative 'abstract_numbering/multilevel_type'
require_relative 'abstract_numbering/numbering_level'
module OoxmlParser
  # This element specifies a set of properties which shall dictate the appearance and
  # behavior of a set of numbered
  # paragraphs in a WordprocessingML document.
  class AbstractNumbering
    # @return [Integer] abstruct numbering id
    attr_accessor :id
    # @return [MultilevelType] myltylevel type
    attr_accessor :multilevel_type
    # @return [Array, NumberingLevel] numbering level data list
    attr_accessor :level_list

    def initialize
      @level_list = []
    end

    # Parse Abstract Numbering data
    # @param [Nokogiri::XML:Element] node with Abstract Numbering data
    # @return [AbstractNumbering] value of Abstract Numbering data
    def self.parse(node)
      abstract_numbering = AbstractNumbering.new

      node.attributes.each do |key, value|
        case key
        when 'abstractNumId'
          abstract_numbering.id = value.value.to_f
        end
      end

      node.xpath('*').each do |numbering_child_node|
        case numbering_child_node.name
        when 'multiLevelType'
          abstract_numbering.multilevel_type = MultilevelType.new(parent: abstract_numbering).parse(numbering_child_node)
        when 'lvl'
          abstract_numbering.level_list << NumberingLevel.parse(numbering_child_node)
        end
      end
      abstract_numbering
    end
  end
end
