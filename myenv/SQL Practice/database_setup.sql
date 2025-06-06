-- Create Authors table if it does not exist
CREATE TABLE IF NOT EXISTS Authors (
    author_id INTEGER PRIMARY KEY,
    author_name TEXT
);

-- Create Books table if it does not exist
CREATE TABLE IF NOT EXISTS Books (
    book_id INTEGER PRIMARY KEY,
    title TEXT,
    author_id INTEGER,
    FOREIGN KEY (author_id) REFERENCES Authors(author_id)
);

-- Insert authors
INSERT OR IGNORE INTO Authors (author_id, author_name) VALUES
(1, 'George Orwell'),
(2, 'Jane Austen'),
(3, 'Harper Lee');

-- Insert books
INSERT OR IGNORE INTO Books (book_id, title, author_id) VALUES
(101, '1984', 1),
(102, 'Pride and Prejudice', 2),
(103, 'To Kill a Mockingbird', 3),
(104, 'Animal Farm', 1),
(105, 'Emma', 2);
