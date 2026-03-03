# AxlsxWorksheetBuilder

Inspired from code originally written by [Samuel Trottier](https://github.com/SamuelTrottier).


## Installation


Install the gem and add to the application's Gemfile by executing:

```bash
bundle add axlsx_worksheet_builder
```

If bundler is not being used to manage dependencies, install the gem by executing:

```bash
gem install axlsx_worksheet_builder
```

## Usage

Add a simple table to an existing axlsx worksheet.

### Iterate through an array of hash
```ruby
authors = [
  {name: "Alice", dob: "1959-01-02", books: [{title: "Book 1"}, {title: "Book 2"}]},
  {name: "Bob", dob: "1962-03-04", books: [{title: "Book A"}]}
]
AxlsxWorksheetBuilder::build(sheet, authors) do |worksheet|
  worksheet.add_column("Author name", property: :name)
  worksheet.add_column("Year of birth") { |author| Date.parse(author[:dob])&.year }
  worksheet.add_column("Number of books") { |author| author[:books]&.size || 0 }
end
```

### Iterate through an array of objects
```ruby
authors = [
  Author.new(name: "Alice", dob: "1959-01-02"),
  Author.new(name: "Bob", dob: "1959-01-02")
]
AxlsxWorksheetBuilder::build(sheet, authors) do |worksheet|
  worksheet.add_column("Author name", property: :name)
  worksheet.add_column("Year of birth") { |author| Date.parse(author.dob)&.year }
end
```

### Iterate through an Array property
```ruby
authors = [{name: "Martin", books: [{title: "Book 1"}, {title: "Book 2"}]}]
AxlsxWorksheetBuilder::build(sheet, authors) do |worksheet|
  worksheet.iterate_through_property(:books)
  worksheet.add_column("Author name", property: :name)
  worksheet.add_column("Number of books") { |author| author.books.count }
  worksheet.add_column("Book title") { |author, book| books.title }
end
```

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/dreeven-oss/axlsx_worksheet_builder.

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).
