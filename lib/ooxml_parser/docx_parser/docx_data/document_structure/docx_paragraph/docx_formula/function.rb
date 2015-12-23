# Function Data
module OoxmlParser
  class Function
    attr_accessor :name, :argument

    def initialize(name = nil, argument = nil)
      @name = name
      @argument = argument
    end
  end
end
