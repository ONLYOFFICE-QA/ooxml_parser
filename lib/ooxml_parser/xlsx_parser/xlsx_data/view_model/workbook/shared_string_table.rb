require_relative 'shared_string_table/string_index'
module OoxmlParser
  # Class for parsing shared string table
  class SharedStringTable < OOXMLDocumentObject
    # @return [Integer] count of shared strings
    attr_reader :count
    # @return [Integer] unique count of shared strings
    attr_reader :unique_count
    # @return [Array<StringIndex>] String Index
    attr_reader :string_indexes

    def initialize(parent: nil)
      @string_indexes = []
      @parent = parent
    end

    # Parse Shared string table file
    # @param file [String] path to file
    # @return [SharedStringTable]
    def parse(file)
      return nil unless File.exist?(file)

      document = Nokogiri::XML(File.open(file))
      node = document.xpath('*').first

      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        when 'uniqueCount'
          @unique_count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'si'
          @string_indexes << StringIndex.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
