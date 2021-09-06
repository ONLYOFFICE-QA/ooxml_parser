# frozen_string_literal: true

module OoxmlParser
  # Parsing `dxfs` tag
  class DifferentialFormattingRecords < OOXMLDocumentObject
    # @return [Array, DifferentialFormattingRecord] list of formatting records
    attr_reader :differential_formatting_records
    # @return [Integer] count of formats
    attr_reader :count

    def initialize(parent: nil)
      @differential_formatting_records = []
      super
    end

    # @return [Array, DifferentialFormattingRecord] accessor
    def [](key)
      @differential_formatting_records[key]
    end

    # Parse DifferentialFormattingRecords data
    # @param [Nokogiri::XML:Element] node with DifferentialFormattingRecords data
    # @return [DifferentialFormattingRecords] value of DifferentialFormattingRecords data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'dxf'
          @differential_formatting_records << DifferentialFormattingRecord.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
