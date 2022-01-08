# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <protectedRange> tag
  class ProtectedRange < OOXMLDocumentObject
    # @return [String] Name of hashing algorithm
    attr_reader :algorithm_name
    # @return [String] Hash value for the password
    attr_reader :hash_value
    # @return [String] Salt value for the password
    attr_reader :salt_value
    # @return [Integer] Number of times the hashing function shall be iteratively run
    attr_reader :spin_count
    # @return [String] Name of protected range
    attr_accessor :name
    # @return [String] Range reference
    attr_reader :reference_sequence

    # Parse ProtectedRange data
    # @param [Nokogiri::XML:Element] node with ProtectedRange data
    # @return [Sheet] value of ProtectedRange
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'algorithmName'
          @algorithm_name = value.value.to_s
        when 'hashValue'
          @hash_value = value.value.to_s
        when 'saltValue'
          @salt_value = value.value.to_s
        when 'spinCount'
          @spin_count = value.value.to_i
        when 'name'
          @name = value.value.to_s
        when 'sqref'
          @reference_sequence = value.value.to_s
        end
      end
      self
    end
  end
end
