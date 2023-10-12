def find_unique_words(sentence1, sentence2):
    words1 = sentence1.split()
    words2 = sentence2.split()
    
    unique_words1 = []
    unique_words2 = []
    
    for word in words1:
        if word not in words2 and word not in unique_words1:
            unique_words1.append(word)
    
    for word in words2:
        if word not in words1 and word not in unique_words2:
            unique_words2.append(word)
    
    return unique_words1, unique_words2

sentence1 = "Sentence Example"
sentence2 = "Sentence Example"
unique_words1, unique_words2 = find_unique_words(sentence1, sentence2)

print("First Sentence:")
print(' '.join(unique_words1))

print("\nSecond Sentence:")
print(' '.join(unique_words2))