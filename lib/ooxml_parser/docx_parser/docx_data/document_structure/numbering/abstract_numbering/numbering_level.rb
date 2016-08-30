require_relative 'numbering_level/level_justification'
require_relative 'numbering_level/level_text'
require_relative 'numbering_level/numbering_format'
require_relative 'numbering_level/start'
module OoxmlParser
  # This element specifies the appearance and behavior of a numbering level
  # within a given abstract numbering
  class NumberingLevel
    # @return [Integer] level id
    attr_accessor :ilvl
    # @return [Start] start data
    attr_accessor :start
    # @return [NumberingFormat] numbering format data
    attr_accessor :numbering_format
    # @return [LevelText] level text
    attr_accessor :text
    # @return [LevelJustification] justification of level
    attr_accessor :justification
    # @return [ParagraphProperties] properties of paragraph
    attr_accessor :paragraph_properties
    # @return [RunProperties] properties of run
    attr_accessor :run_properties

    # Parse Numbering Level data
    # @param [Nokogiri::XML:Element] node with Numbering Level data
    # @return [NumberingLevel] value of Numbering Level data
    def self.parse(node)
      level = NumberingLevel.new

      node.attributes.each do |key, value|
        case key
        when 'ilvl'
          level.ilvl = value.value.to_f
        end
      end

      node.xpath('*').each do |num_level_child|
        case num_level_child.name
        when 'start'
          level.start = Start.parse(num_level_child)
        when 'numFmt'
          level.numbering_format = NumberingFormat.parse(num_level_child)
        when 'lvlText'
          level.text = LevelText.parse(num_level_child)
        when 'lvlJc'
          level.justification = LevelJustification.parse(num_level_child)
        when 'pPr'
          level.paragraph_properties = ParagraphProperties.parse(num_level_child)
        when 'rPr'
          level.run_properties = RunProperties.new(parent: level).parse(num_level_child)
        end
      end

      level
    end
  end
end
