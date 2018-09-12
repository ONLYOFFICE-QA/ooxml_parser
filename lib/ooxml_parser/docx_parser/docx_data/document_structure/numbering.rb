require_relative 'numbering/abstract_numbering'
require_relative 'numbering/numbering_definition'

module OoxmlParser
  class Numbering < OOXMLDocumentObject
    # @return [Array, AbstractNumbering] abstract numbering list
    attr_accessor :abstract_numbering_list
    # @return [Array, NumberingDefinition] numbering definition list
    attr_accessor :numbering_definition_list

    def initialize(parent: nil)
      @abstract_numbering_list = []
      @numbering_definition_list = []
      @parent = parent
    end

    def properties_by_num_id(num_id)
      abstract_num_id = nil
      @numbering_definition_list.each do |num_def|
        next unless num_id == num_def.id

        abstract_num_id = num_def.abstract_numbering_id.value
        break
      end
      return nil if abstract_num_id.nil?

      @abstract_numbering_list.each do |abstract_num_item|
        return abstract_num_item if abstract_num_id == abstract_num_item.id
      end
    end

    def parse
      numbering_xml = OOXMLDocumentObject.path_to_folder + 'word/numbering.xml'
      return nil unless File.exist?(numbering_xml)

      node = Nokogiri::XML(File.open(numbering_xml), 'r:UTF-8')
      node.xpath('w:numbering/*').each do |numbering_child_node|
        case numbering_child_node.name
        when 'abstractNum'
          @abstract_numbering_list << AbstractNumbering.new(parent: self).parse(numbering_child_node)
        when 'num'
          @numbering_definition_list << NumberingDefinition.new(parent: self).parse(numbering_child_node)
        end
      end
      self
    end
  end
end
