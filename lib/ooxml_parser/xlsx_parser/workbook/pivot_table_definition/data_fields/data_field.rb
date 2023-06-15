# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <dataField> tag
  class DataField < OOXMLDocumentObject
    # @return [Integer] index of the base field for the ShowDataAs calculation
    attr_accessor :base_field
    # @return [Integer] index of the base item for the ShowDataAs calculation
    attr_accessor :base_item
    # @return [Integer] index of the field in the pivotCacheRecords
    attr_accessor :field
    # @return [String] name of the data field
    attr_reader :name
    # @return [Integer] index of the number format applied to data field
    attr_accessor :number_format_id
    # @return [ExtensionList] list of extensions
    attr_accessor :extension_list

    # Parse `<dataField>` tag
    # # @param [Nokogiri::XML:Element] node with DataField data
    # @return [DataField]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'baseField'
          @base_field = value.value.to_i
        when 'baseItem'
          @base_item = value.value.to_i
        when 'fld'
          @field = value.value.to_i
        when 'name'
          @name = value.value.to_s
        when 'numFmtId'
          @number_format_id = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'extLst'
          @extension_list = ExtensionList.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
