const express = require('express');
const fs = require('fs');
const bodyParser = require('body-parser');

const app = express();
app.use(bodyParser.json());

app.get('/user/:userId', (req, res) => {
  const userId = req.params.userId;
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const user = data.users.find((user) => user.id === parseInt(userId));
  if (user) {
    res.send(user);
  } else {
    res.status(404).send('User not found');
  }
});

app.get('/post/:postId', (req, res) => {
  const { postId } = req.params;
  const { posts } = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const post = posts.find((post) => post.id === parseInt(postId));

  if (post) {
    res.send(post);
  } else {
    res.status(404).send('Post not found');
  }
});

app.listen(3000, () => {
  console.log('Server running on 3000');
});

app.get('/posts/:startDate/:endDate', (req, res) => {
  const { startDate, endDate } = req.params;
  const { posts } = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const filteredPosts = posts.filter((post) => {
    const postDate = new Date(post.last_update);
    return postDate >= new Date(startDate) && postDate <= new Date(endDate);
  });

  if (filteredPosts.length > 0) {
    res.send(filteredPosts);
  } else {
    res.status(404).send('No posts found between specified dates');
  }
});

app.post('/user/:userId/email', (req, res) => {
  const userId = req.params.userId;
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

  const user = data.users.find((user) => user.id === parseInt(userId));

  if (user) {
    user.email = req.body.email;
    fs.writeFileSync('data.json', JSON.stringify(data));
    res.send('Email updated');
  } else {
    res.status(404).send('User not found');
  }
});

app.put('/user/:userId/post', (req, res) => {
  const { userId } = req.params;
  const { users, posts } = JSON.parse(fs.readFileSync('data.json', 'utf8'));
  const { title, body } = req.body;

  const userIndex = users.findIndex((user) => user.id === parseInt(userId));

  if (userIndex !== -1) {
    const newPost = {
      id: posts.length + 1,
      userId: parseInt(userId),
      title,
      body,
      date: new Date().toISOString(),
      last_update: new Date().toISOString(),
    };
    posts.push(newPost);
    fs.writeFileSync('data.json', JSON.stringify({ users, posts }));
    res.send('Post created');
  } else {
    res.status(404).send('User not found');
  }
});
